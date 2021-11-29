VERSION 5.00
Begin VB.Form frm132 
   Caption         =   "Data Dulang"
   ClientHeight    =   11325
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   21030
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
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11325
   ScaleWidth      =   21030
   WindowState     =   2  'Maximized
   Begin VB.Timer Tmr2 
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Dulang"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   9720
      TabIndex        =   4
      Top             =   240
      Width           =   8895
      Begin VB.CommandButton CMD5 
         Caption         =   "Simpan Data"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   4680
         MouseIcon       =   "frm132.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "frm132.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1320
         Width           =   2775
      End
      Begin VB.CheckBox CB1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   480
         TabIndex        =   8
         Top             =   375
         Width           =   200
      End
      Begin VB.TextBox TB1 
         BackColor       =   &H8000000A&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   3960
         TabIndex        =   7
         Text            =   "TB1"
         Top             =   600
         Width           =   4395
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Scanner Mode"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   765
         TabIndex        =   9
         Top             =   345
         Width           =   6930
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "No. Siri Produk :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   5
         Top             =   600
         Width           =   3375
      End
   End
   Begin VB.CommandButton CMD2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tukar Dulang"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   240
      MouseIcon       =   "frm132.frx":28D4
      MousePointer    =   99  'Custom
      Picture         =   "frm132.frx":2BDE
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label L2_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L2_Text"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   360
      TabIndex        =   11
      Top             =   4680
      Width           =   18375
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   $"frm132.frx":93E8
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   360
      TabIndex        =   6
      Top             =   3480
      Width           =   14295
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   ">> Sila pilih DULANG. Semua barang yang discan dari menu ini akan dipindahkan ke dalam dulang yang dipilih."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   2880
      Width           =   14295
   End
   Begin VB.Label L1_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "L1_Text"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6360
      TabIndex        =   2
      Top             =   240
      Width           =   3375
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Dulang :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2760
      TabIndex        =   1
      Top             =   240
      Width           =   3375
   End
End
Attribute VB_Name = "frm132"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMD2_Click()
'on error resume next
Call frm133_setting
frm133.Show vbModal
End Sub

Private Sub CMD5_Click()
'On Error Resume Next
If frm132.TB1 = vbNullString Then
    
    MsgBox "Sila masukkan no.siri produk.", vbExclamation, "Info"
    
    frm132.TB1.SetFocus
    
    Exit Sub
    
End If
If frm132.L1_Text = vbNullString Then
    
    MsgBox "Sila buat pilihan dulang.", vbExclamation, "Info"
    
    frm132.TB1.SetFocus
    
    Exit Sub
    
End If

If InStr(1, frm132.TB1, "'") <> 0 Then
    MsgBox "No. Siri Produk Mengandungi Simbol Yang Tidak Sah , ['].", vbInformation, "Info"
    
    frm132.TB1 = vbNullString
    frm132.TB1.SetFocus
    
    Exit Sub
End If

Call frm132_tukar_dulang
End Sub

Private Sub L1_Text_Change()
'on error resume next
If frm132.L1_Text = vbNullString Then
    frm132.Frame1.Visible = False
Else
    frm132.Frame1.Visible = True
End If
End Sub
Private Sub TB1_Change()
'on error resume next
If frm132.CB1 = 1 And frm132.TB1 <> vbNullString Then
    frm132.Tmr2.Enabled = False
    frm132.Tmr2.Enabled = True
    frm132.Tmr2.Interval = 100
End If
End Sub

Private Sub Tmr2_Timer()
'On Error Resume Next
DATA_UDPATE = 0

If frm132.CB1 = 1 And frm132.TB1 <> vbNullString And frm132.Tmr2.Enabled = True Then

    If frm132.Tmr2.Interval = 100 Then
        If InStr(1, frm132.TB1, "'") <> 0 Then
            MsgBox "No. Siri Produk Mengandungi Simbol Yang Tidak Sah , ['].", vbInformation, "Info"
            
            frm132.TB1 = vbNullString
            frm132.TB1.SetFocus
            
            Exit Sub
        End If
        
        Call frm132_tukar_dulang
        
    End If
End If
End Sub
