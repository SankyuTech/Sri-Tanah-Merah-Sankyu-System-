VERSION 5.00
Begin VB.Form Frm97 
   Caption         =   "E-mail Promosi"
   ClientHeight    =   13035
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   23760
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Frm97.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   13035
   ScaleWidth      =   23760
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Pic2 
      BorderStyle     =   0  'None
      Height          =   9855
      Left            =   6000
      ScaleHeight     =   9855
      ScaleWidth      =   10215
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   10215
      Begin VB.CommandButton CMD3 
         BackColor       =   &H000080FF&
         Caption         =   "Batal"
         Height          =   345
         Left            =   5160
         MaskColor       =   &H00400000&
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   9360
         Width           =   1545
      End
      Begin VB.CommandButton CMD2 
         BackColor       =   &H000080FF&
         Caption         =   "Hantar Email"
         Height          =   345
         Left            =   3480
         MaskColor       =   &H00400000&
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   9360
         Width           =   1545
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Butiran E-mail Promosi"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   360
         TabIndex        =   15
         Top             =   840
         Width           =   8055
      End
      Begin VB.Label L7_Text 
         Caption         =   "L7_Text"
         ForeColor       =   &H00000000&
         Height          =   8055
         Left            =   240
         TabIndex        =   6
         Top             =   1200
         Width           =   9495
      End
      Begin VB.Label L6_Text 
         Caption         =   "L6_Text"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1320
         TabIndex        =   5
         Top             =   360
         Width           =   8895
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Subjek   :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   360
         TabIndex        =   16
         Top             =   360
         Width           =   8055
      End
   End
   Begin VB.PictureBox Pic1 
      BorderStyle     =   0  'None
      Height          =   10095
      Left            =   8160
      ScaleHeight     =   10095
      ScaleWidth      =   10215
      TabIndex        =   7
      Top             =   2400
      Visible         =   0   'False
      Width           =   10215
      Begin VB.TextBox TB2 
         Height          =   7335
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   10
         Text            =   "Frm97.frx":0ECA
         Top             =   2160
         Width           =   9735
      End
      Begin VB.CommandButton CMD1 
         BackColor       =   &H000080FF&
         Caption         =   "Simpan Data"
         Height          =   345
         Left            =   3960
         MaskColor       =   &H00400000&
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   9600
         Width           =   1545
      End
      Begin VB.TextBox TB1 
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Text            =   "TB1"
         Top             =   480
         Width           =   9735
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Kepada Tuan/Puan/Encik/Cik *************"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   1800
         Width           =   9615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Assalamualaikum Dan Selamat Sejahtera."
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Top             =   1440
         Width           =   9615
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Subjek E-mail"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   240
         Width           =   8055
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Butiran E-mail Promosi"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   1080
         Width           =   8055
      End
   End
   Begin VB.Label L5_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Hantar E-mail"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      MouseIcon       =   "Frm97.frx":0ECE
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label L4_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Hasilkan Isi E-mail Promosi"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      MouseIcon       =   "Frm97.frx":11D8
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   0
      Width           =   2895
   End
End
Attribute VB_Name = "Frm97"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMD1_Click()
'on error resume next
Dim Err(10)

DATA_WRITE = 0 '0 : Tiada Data Disimpan , 1 : Data Telah Disimpan
If Frm97.TB1 = vbNullString Then
    x = x + 1
    Err(x) = "Sila masukkan subjek."
End If
If Frm97.TB2 = vbNullString Then
    x = x + 1
    Err(x) = "Sila masukkan isi kandungan promosi."
End If

If x <> 0 Then
    Frm6.Show
    Frm6.Pic1.Cls
    For Y = 1 To x
        Frm6.Pic1.Print Y & " - " & Err(Y)
    Next Y
    Exit Sub
Else
    Note = "Adakah Anda Ingin Simpan Data Ini ?"
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbNo Then
        Exit Sub
    End If
    If Answer = vbYes Then
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 45_email_promosi where default_setting='" & "default" & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Frm97.TB1 <> vbNullString Then
                rs!subjek = Frm97.TB1 'Subject
            Else
                rs!subjek = Null 'Subject
            End If
            If Frm97.TB2 <> vbNullString Then
                rs!Body = Frm97.TB2 'Body
            Else
                rs!Body = Null 'Body
            End If
            rs.Update
        End If
        
        rs.Close
        Set rs = Nothing
        
        MsgBox "Maklumat promosi telah berjaya disimpan.", vbInformation, "Info"
    End If
End If
End Sub
Private Sub CMD2_Click()
'on error resume next
Call check_internet_connection
End Sub
Private Sub CMD3_Click()
'on error resume next
Frm97.Pic2.Visible = False
End Sub
Private Sub L4_Text_Click()
'on error resume next
If Frm97.Pic1.Visible = False Then
    Call Frm97_initial_setting
    Call Frm97_Call_Promosi
    
    Frm97.Pic1.Visible = True
Else
    Frm97.Pic1.Visible = False
End If
End Sub
Private Sub L5_Text_Click()
'on error resume next
If Frm97.Pic2.Visible = False Then
    Call Frm97_initial_setting
    Call Frm97_Call_Promosi2
    
    Frm97.Pic2.Visible = True
Else
    Frm97.Pic2.Visible = False
End If
End Sub
