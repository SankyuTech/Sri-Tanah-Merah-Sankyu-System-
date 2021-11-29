VERSION 5.00
Begin VB.Form Frm114 
   Caption         =   "Kalkulator Trade In"
   ClientHeight    =   13035
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   23880
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
   MDIChild        =   -1  'True
   ScaleHeight     =   13035
   ScaleWidth      =   23880
   WindowState     =   2  'Maximized
   Begin VB.TextBox TB1 
      BackColor       =   &H80000001&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   645
      Left            =   4920
      TabIndex        =   2
      Text            =   "TB1"
      Top             =   400
      Width           =   2775
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tael :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      TabIndex        =   20
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Public :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      TabIndex        =   19
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "SA :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      TabIndex        =   18
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Harga Besar (Spot Price) :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   17
      Top             =   480
      Width           =   4215
   End
   Begin VB.Label L14_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "L14_Text"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10095
      Left            =   11280
      TabIndex        =   16
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label L13_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah Untung"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11160
      TabIndex        =   15
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Label L12_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "L12_Text"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9975
      Left            =   9000
      TabIndex        =   14
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label L11_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Untung Atas Mutu"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8880
      TabIndex        =   13
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Label L10_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "L10_Text"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9975
      Left            =   6600
      TabIndex        =   12
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label L9_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Untung Atas Tolak"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6480
      TabIndex        =   11
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Label L7_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "L7_Text"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   9975
      Left            =   2040
      TabIndex        =   10
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label L8_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SA"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   9
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Label L6_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "L6_Text"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   9120
      TabIndex        =   8
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Harga Beli"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   7
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Purity"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   195
      TabIndex        =   6
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label L5_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "L5_Text"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9975
      Left            =   4320
      TabIndex        =   5
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label L3_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "L3_Text"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   9120
      TabIndex        =   4
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label L2_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "L2_Text"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   9120
      TabIndex        =   3
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label L4_Text 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "L4_Text"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9975
      Left            =   240
      TabIndex        =   1
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label L1_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Keluar"
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
      Left            =   19800
      MouseIcon       =   "Frm114.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Frm114"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'On Error Resume Next
If MDI_frm1.L20_Text = "Semua cawangan" Then
    LM_CAWANGAN = "HQ"
Else
    LM_CAWANGAN = MDI_frm1.L20_Text
End If

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from default_setting where Default1='" & LM_CAWANGAN & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    
    If Not IsNull(rs!tael) Then Frm114.L3_Text = rs!tael 'Tael
    If Not IsNull(rs!public) Then Frm114.L6_Text = rs!public 'Public
    If Not IsNull(rs!sa) Then Frm114.L2_Text = rs!sa 'SA
                
End If

rs.Close
Set rs = Nothing

Frm114.L4_Text.Visible = False
Frm114.L5_Text.Visible = False
Frm114.L7_Text.Visible = False
Frm114.L9_Text.Visible = False
Frm114.L10_Text.Visible = False
Frm114.L11_Text.Visible = False
Frm114.L12_Text.Visible = False
Frm114.L13_Text.Visible = False
Frm114.L14_Text.Visible = False

If MDI_frm1.L4_Text = "Admin" Or MDI_frm1.L4_Text = "HQ" Or user_level = "Developer" Then

    Frm114.L9_Text.Visible = True
    Frm114.L10_Text.Visible = True
    Frm114.L11_Text.Visible = True
    Frm114.L12_Text.Visible = True
    Frm114.L13_Text.Visible = True
    Frm114.L14_Text.Visible = True
    
Else

    Frm114.L9_Text.Visible = False
    Frm114.L10_Text.Visible = False
    Frm114.L11_Text.Visible = False
    Frm114.L12_Text.Visible = False
    Frm114.L13_Text.Visible = False
    Frm114.L14_Text.Visible = False
    
End If

Frm114.Picture = MDI_frm1.Picture

Frm114.L4_Text = vbNullString
Frm114.L5_Text = vbNullString

Frm114.TB1 = MDI_frm1.L16_Text
End Sub
Private Sub L1_Text_Click()
Unload Frm114
End Sub

Private Sub TB1_Change()
If IsNumeric(Frm114.TB1) And Len(Frm114.TB1) = 4 Then
    MDI_frm1.L16_Text = Frm114.TB1
    Call frm114_kalkulator_ti
    
    Frm114.L4_Text.Visible = True
    Frm114.L5_Text.Visible = True
    Frm114.L7_Text.Visible = True
    Frm114.L10_Text.Visible = True
    Frm114.L12_Text.Visible = True
    Frm114.L14_Text.Visible = True
    
    If MDI_frm1.L4_Text = "Admin" Or MDI_frm1.L4_Text = "HQ" Or user_level = "Developer" Then
    
        Frm114.L9_Text.Visible = True
        Frm114.L10_Text.Visible = True
        Frm114.L11_Text.Visible = True
        Frm114.L12_Text.Visible = True
        Frm114.L13_Text.Visible = True
        Frm114.L14_Text.Visible = True
        
    Else
    
        Frm114.L9_Text.Visible = False
        Frm114.L10_Text.Visible = False
        Frm114.L11_Text.Visible = False
        Frm114.L12_Text.Visible = False
        Frm114.L13_Text.Visible = False
        Frm114.L14_Text.Visible = False
        
    End If
    
Else
    MDI_frm1.L16_Text = 0
    Frm114.L4_Text.Visible = False
    Frm114.L5_Text.Visible = False
    Frm114.L7_Text.Visible = False
    Frm114.L10_Text.Visible = False
    Frm114.L12_Text.Visible = False
    Frm114.L14_Text.Visible = False
End If
End Sub
