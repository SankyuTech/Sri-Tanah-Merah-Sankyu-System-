VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm100 
   Caption         =   "Kemasukkan dan pengeluaran tunai kedai"
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
   Icon            =   "Frm100.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   13035
   ScaleWidth      =   23760
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Pic2 
      BorderStyle     =   0  'None
      Height          =   11295
      Left            =   10920
      ScaleHeight     =   11295
      ScaleWidth      =   18735
      TabIndex        =   1
      Top             =   1080
      Visible         =   0   'False
      Width           =   18735
      Begin VB.CommandButton CMD20 
         Caption         =   "Paparan Seterusnya"
         Height          =   375
         Left            =   6240
         MouseIcon       =   "Frm100.frx":0ECA
         MousePointer    =   99  'Custom
         TabIndex        =   56
         Top             =   9600
         Width           =   1935
      End
      Begin VB.CommandButton CMD19 
         Caption         =   "Paparan Sebelum"
         Height          =   375
         Left            =   4200
         MouseIcon       =   "Frm100.frx":11D4
         MousePointer    =   99  'Custom
         TabIndex        =   55
         Top             =   9600
         Width           =   1935
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   9075
         Left            =   120
         TabIndex        =   40
         ToolTipText     =   "Sila ""Double Click"" untuk menu seterusnya."
         Top             =   480
         Width           =   8085
         _ExtentX        =   14261
         _ExtentY        =   16007
         _Version        =   393216
         Rows            =   1
         Cols            =   0
         FixedCols       =   0
         BackColor       =   16777215
         ForeColor       =   0
         BackColorFixed  =   12648384
         BackColorSel    =   16777215
         ForeColorSel    =   16711680
         BackColorBkg    =   16777215
         GridColor       =   0
         WordWrap        =   -1  'True
         ScrollTrack     =   -1  'True
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label L13_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L13_Text"
         ForeColor       =   &H00000000&
         Height          =   9015
         Left            =   8280
         TabIndex        =   39
         Top             =   600
         Width           =   8000
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah ambilan tunai       : RM"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   10080
         Width           =   2775
      End
      Begin VB.Label L12_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3000
         TabIndex        =   34
         Top             =   10080
         Width           =   1575
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Paparan Muka   :"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4680
         TabIndex        =   33
         Top             =   10080
         Width           =   1695
      End
      Begin VB.Label L10_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L10_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   6240
         TabIndex        =   32
         Top             =   10080
         Width           =   900
      End
      Begin VB.Label L11_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3000
         TabIndex        =   4
         Top             =   9840
         Width           =   1575
      End
      Begin VB.Label Label58 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah kemasukkan tunai : RM"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   9840
         Width           =   2775
      End
      Begin VB.Label L7_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "Rekod"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   11295
      End
   End
   Begin VB.PictureBox Pic3 
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   9000
      ScaleHeight     =   2535
      ScaleWidth      =   9705
      TabIndex        =   24
      Top             =   2880
      Visible         =   0   'False
      Width           =   9705
      Begin VB.CommandButton CMD5 
         Caption         =   "Batal"
         Height          =   375
         Left            =   4680
         MouseIcon       =   "Frm100.frx":14DE
         MousePointer    =   99  'Custom
         TabIndex        =   54
         Top             =   1920
         Width           =   1935
      End
      Begin VB.CommandButton CMD4 
         Caption         =   "Rekod"
         Height          =   375
         Left            =   2640
         MouseIcon       =   "Frm100.frx":17E8
         MousePointer    =   99  'Custom
         TabIndex        =   53
         Top             =   1920
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   360
         Left            =   2280
         TabIndex        =   25
         Top             =   885
         Width           =   7005
         _ExtentX        =   12356
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   16744576
         Format          =   142475264
         CurrentDate     =   41561
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   360
         Left            =   2280
         TabIndex        =   26
         Top             =   1245
         Width           =   7005
         _ExtentX        =   12356
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   16744576
         Format          =   142475264
         CurrentDate     =   41561
      End
      Begin VB.Shape Shape2 
         Height          =   1455
         Left            =   240
         Top             =   285
         Width           =   9255
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Sila buat pilihan tarikh rekod kemasukkan atau pengeluaran tunai."
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   360
         TabIndex        =   31
         Top             =   480
         Width           =   8610
      End
      Begin VB.Label Label63 
         BackStyle       =   0  'Transparent
         Caption         =   "Tarikh Akhir "
         Height          =   255
         Left            =   315
         TabIndex        =   30
         Top             =   1290
         Width           =   2895
      End
      Begin VB.Label Label62 
         BackStyle       =   0  'Transparent
         Caption         =   "Tarikh Mula "
         Height          =   255
         Left            =   315
         TabIndex        =   29
         Top             =   930
         Width           =   2535
      End
      Begin VB.Label L9_Text 
         Caption         =   "L9_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   8640
         TabIndex        =   28
         Top             =   1920
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Label L8_Text 
         Caption         =   "L8_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   7800
         TabIndex        =   27
         Top             =   1920
         Visible         =   0   'False
         Width           =   660
      End
   End
   Begin VB.Timer Tmr1 
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox Pic1 
      BorderStyle     =   0  'None
      Height          =   8895
      Left            =   2160
      ScaleHeight     =   8895
      ScaleWidth      =   10095
      TabIndex        =   5
      Top             =   840
      Visible         =   0   'False
      Width           =   10095
      Begin VB.CommandButton CMD3 
         Caption         =   "Batal"
         Height          =   375
         Left            =   4200
         MouseIcon       =   "Frm100.frx":1AF2
         MousePointer    =   99  'Custom
         TabIndex        =   52
         Top             =   6120
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CommandButton CMD2 
         Caption         =   "Simpan Data"
         Height          =   375
         Left            =   2040
         MouseIcon       =   "Frm100.frx":1DFC
         MousePointer    =   99  'Custom
         TabIndex        =   51
         Top             =   6120
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CommandButton CMD1 
         Caption         =   "Simpan Data"
         Height          =   375
         Left            =   3120
         MouseIcon       =   "Frm100.frx":2106
         MousePointer    =   99  'Custom
         TabIndex        =   50
         Top             =   6120
         Width           =   1935
      End
      Begin VB.TextBox TB5 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   2130
         TabIndex        =   47
         Text            =   "TB5"
         Top             =   2280
         Width           =   4245
      End
      Begin VB.TextBox TB4 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   2130
         TabIndex        =   44
         Text            =   "TB4"
         Top             =   1920
         Width           =   4245
      End
      Begin VB.TextBox TB3 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   2130
         TabIndex        =   41
         Text            =   "TB3"
         Top             =   1560
         Width           =   4245
      End
      Begin VB.TextBox TB2 
         BackColor       =   &H00FFFFFF&
         Height          =   1920
         Left            =   2130
         MultiLine       =   -1  'True
         TabIndex        =   36
         Text            =   "Frm100.frx":2410
         Top             =   3960
         Width           =   4245
      End
      Begin VB.ComboBox CBB1 
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         Height          =   360
         Left            =   2130
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   3600
         Width           =   4245
      End
      Begin VB.CheckBox CB2 
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
         Left            =   3120
         TabIndex        =   13
         Top             =   1080
         Width           =   200
      End
      Begin VB.CheckBox CB1 
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
         Left            =   645
         TabIndex        =   10
         Top             =   1065
         Width           =   200
      End
      Begin VB.TextBox TB1 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   2130
         TabIndex        =   9
         Text            =   "TB1"
         Top             =   2880
         Width           =   4245
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   360
         Left            =   2130
         TabIndex        =   19
         Top             =   3240
         Width           =   4245
         _ExtentX        =   7488
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   16744576
         Format          =   415301632
         CurrentDate     =   41561
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   2040
         TabIndex        =   49
         Top             =   2310
         Width           =   120
      End
      Begin VB.Label L16_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "No. Telefon"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   480
         TabIndex        =   48
         Top             =   2310
         Width           =   1545
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   2040
         TabIndex        =   46
         Top             =   1950
         Width           =   120
      End
      Begin VB.Label L15_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "Kad Pengenalan"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   480
         TabIndex        =   45
         Top             =   1950
         Width           =   1545
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   2040
         TabIndex        =   43
         Top             =   1590
         Width           =   120
      End
      Begin VB.Label L14_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   480
         TabIndex        =   42
         Top             =   1590
         Width           =   1545
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   2040
         TabIndex        =   38
         Top             =   3990
         Width           =   120
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   480
         TabIndex        =   37
         Top             =   3990
         Width           =   1545
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   2040
         TabIndex        =   23
         Top             =   3600
         Width           =   120
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   2040
         TabIndex        =   22
         Top             =   3240
         Width           =   120
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Pekerja *"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   480
         TabIndex        =   21
         Top             =   3600
         Width           =   2295
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "Tarikh *"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   480
         TabIndex        =   20
         Top             =   3240
         Width           =   2385
      End
      Begin VB.Shape Shape1 
         Height          =   855
         Left            =   360
         Top             =   600
         Width           =   4575
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Jenis"
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
         Height          =   300
         Left            =   480
         TabIndex        =   17
         Top             =   720
         Width           =   825
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Sila masukkan data terperinci bagi urusan kemasukkan tunai ke dalam kedai atau pengeluaran duit dari kedai."
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   360
         TabIndex        =   16
         Top             =   120
         Width           =   6075
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   11295
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   11295
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Kemasukkan duit                Pengeluaran duit"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   960
         TabIndex        =   12
         Top             =   1035
         Width           =   4425
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah *        RM"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   480
         TabIndex        =   11
         Top             =   2910
         Width           =   1545
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   2040
         TabIndex        =   7
         Top             =   2910
         Width           =   120
      End
      Begin VB.Label L5_Text 
         Caption         =   "L5_Text"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   8400
         TabIndex        =   6
         Top             =   1320
         Visible         =   0   'False
         Width           =   975
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Senarai rekod kemasukkan dan pengeluaran tunai"
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
      Left            =   3720
      MouseIcon       =   "Frm100.frx":2414
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   120
      Width           =   4575
   End
   Begin VB.Label L3_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Kemasukkan atau pengeluaran"
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
      Left            =   0
      MouseIcon       =   "Frm100.frx":271E
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   120
      Width           =   3615
   End
   Begin VB.Menu frm100_pm_menu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu frm100_sm_edit_data 
         Caption         =   "Lihat / edit data"
      End
      Begin VB.Menu Frm100_SM_cetak_voucher 
         Caption         =   "Cetak voucher pengeluaran duit"
      End
      Begin VB.Menu Frm100_SM_terperinci 
         Caption         =   "Lihat data terperinci"
      End
      Begin VB.Menu frm100_sm_padam 
         Caption         =   "Padam data"
      End
   End
End
Attribute VB_Name = "Frm100"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CB1_Click()
'On Error Resume Next
If Frm100.CB1 = 1 Then
    Frm100.CB2 = 0
    
    Frm100.L14_Text = "Nama"
    Frm100.L15_Text = "Kad Pengenalan"
    Frm100.L16_Text = "No. Telefon"
    
    Frm100.TB3 = vbNullString
    Frm100.TB4 = vbNullString
    Frm100.TB5 = vbNullString
    
    Frm100.TB3.Locked = True
    Frm100.TB4.Locked = True
    Frm100.TB5.Locked = True
    
    Frm100.TB3.BackColor = &H8000000A
    Frm100.TB4.BackColor = &H8000000A
    Frm100.TB5.BackColor = &H8000000A
End If
End Sub
Private Sub CB2_Click()
'On Error Resume Next
If Frm100.CB2 = 1 Then
    Frm100.CB1 = 0
    
    Frm100.L14_Text = "Nama *"
    Frm100.L15_Text = "Kad Pengenalan *"
    Frm100.L16_Text = "No. Telefon *"
    
    Frm100.TB3.Locked = False
    Frm100.TB4.Locked = False
    Frm100.TB5.Locked = False
    
    Frm100.TB3.BackColor = &HFFFFFF
    Frm100.TB4.BackColor = &HFFFFFF
    Frm100.TB5.BackColor = &HFFFFFF
End If
End Sub
Private Sub CMD1_Click()
'On Error Resume Next
Dim Err(6)
DATA_SAVE = 0

Frm100_LM_EMP_NAME = vbNullString
Frm100_LM_EMP_NO = vbNullString
Frm100_LM_VOUCHER = 1

If Frm100.CB1 = 0 And Frm100.CB2 = 0 Then
    x = x + 1
    Err(x) = "Sila pilih samada [Kemasukan Duit] atau [Pengeluaran Duit]."
End If
If Frm100.CBB1 = vbNullString Then
    x = x + 1
    Err(x) = "Sila pilih [Nama Pekerja]."
End If
If Frm100.TB1 = vbNullString Or (Frm100.TB1 <> vbNullString And Not IsNumeric(Frm100.TB1)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Jumlah]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm100.CB2 = 1 Then

    If Frm100.TB3 = vbNullString Then
        x = x + 1
        Err(x) = "Sila masukkan [Nama]."
    End If
    If Frm100.TB4 = vbNullString Then
        x = x + 1
        Err(x) = "Sila masukkan [No. Kad Pengenalan]."
    End If
    If Frm100.TB5 = vbNullString Then
        x = x + 1
        Err(x) = "Sila masukkan [No. Telefon]."
    End If
    
End If

If x <> 0 Then
    Frm6.Show
    Frm6.Pic1.Cls
    For Y = 1 To x
        Frm6.Pic1.Print Y & " - " & Err(Y)
    Next Y
    Exit Sub
Else
    Note = "Adakah anda ingin simpan data ini ?"
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbNo Then
        Exit Sub
    End If
    If Answer = vbYes Then
    
        If Frm100.CB2 = 1 Then
        
        '###Carian No. Voucher Terbaru### - Start
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from default_setting", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                If rs!Default1 = "Default" Then
                    If IsNumeric(rs!voucher_duit) Then
                        Frm100_LM_VOUCHER = rs!voucher_duit 'No. voucher
                    Else
                        Frm100_LM_VOUCHER = 1
                    End If
                Else
                    Frm100_LM_VOUCHER = 1
                End If
            Else
                Frm100_LM_VOUCHER = 1
            End If
            
            rs.Close
            Set rs = Nothing
        '###Carian No. Voucher Terbaru### - End
        
        '###Periksa No. Voucher### - Start
re_check_voucher:
        
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from 47_account_close where no_voucher='" & "VOU-" & Format(Frm100_LM_VOUCHER, "000000") & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                
                Frm100_LM_VOUCHER = Frm100_LM_VOUCHER + 1
                
                rs.Close
                Set rs = Nothing
                
                GoTo re_check_voucher:
                
            End If
            
            rs.Close
            Set rs = Nothing
        '###Periksa No. Voucher### - End
        End If
    
        If Frm100.CBB1 <> vbNullString Then
            Frm100_LM_EMP_NAME = Split(Frm100.CBB1, "  |  ")(0)
            Frm100_LM_EMP_NO = Split(Frm100.CBB1, "  |  ")(1)
        End If

        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 47_account_close", cn, adOpenKeyset, adLockOptimistic

        rs.AddNew
        If Frm100.CB1 = 1 Then '0 : Kemasukkan Duit , 1 : Pengeluaran Duit
            rs!jenis = 0
        ElseIf Frm100.CB2 = 1 Then
            rs!jenis = 1
        End If
        rs!tarikh = Frm100.DTPicker1 'Tarikh
        If Frm100.TB1 <> vbNullString Then 'Jumlah
            rs!jumlah = Format(Frm100.TB1, "0.00")
        Else
            rs!jumlah = Format(0, "0.00")
        End If
        If Frm100.TB2 <> vbNullString Then 'Remarks
            rs!remarks = Frm100.TB2
        Else
            rs!remarks = Null
        End If
        If Frm100.CBB1 <> vbNullString Then 'Nama Pekerja
            rs!staff_name = Frm100_LM_EMP_NAME
        Else
            rs!staff_name = Null
        End If
        If Frm100.CBB1 <> vbNullString Then 'Nama Pekerja
            rs!staff_id = Frm100_LM_EMP_NO
        Else
            rs!staff_id = Null
        End If
        If Frm100.TB3 <> vbNullString Then 'Nama
            rs!Nama = UCase(Frm100.TB3)
        Else
            rs!Nama = Null
        End If
        If Frm100.TB4 <> vbNullString Then 'No. kad pengenalan
            rs!no_ic = UCase(Frm100.TB4)
        Else
            rs!no_ic = Null
        End If
        If Frm100.TB5 <> vbNullString Then 'No. telefon
            rs!no_tel = UCase(Frm100.TB5)
        Else
            rs!no_tel = Null
        End If
        If Frm100.CB2 = 1 Then
            rs!no_voucher = "VOU-" & Format(Frm100_LM_VOUCHER, "000000")
        Else
            rs!no_voucher = Null
        End If
        rs!Status = 1 '0 : Tidak Aktif , 1 : Aktif
        rs!write_timestamp = Now
        rs.Update
        
        rs.Close
        Set rs = Nothing

        user = MDI_frm1.L3_Text
        If Frm100.CB1 = 1 Then LogAct_Memory = "[" & user & "] Kemasukkan tunai ke dalam kedai [" & Format(Frm100.TB1, "#,##0.00") & "]."
        If Frm100.CB2 = 1 Then LogAct_Memory = "[" & user & "] Pengeluaran tunai dari kedai [" & "VOU-" & Format(Frm100_LM_VOUCHER, "000000") & "]."
        LogDate_Memory = DateTime.Date & " " & DateTime.Time$
        Call UpdateLog_Database
        
        If Frm100.CB2 = 1 Then
        
'###Carian No. Voucher Terbaru### - Start
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from default_setting where default1='" & "default" & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                rs!voucher_duit = Frm100_LM_VOUCHER + 1
                rs.Update
            End If
            
            rs.Close
            Set rs = Nothing
'###Carian No. Voucher Terbaru### - End
        
        End If
        
        Frm100.TB1 = vbNullString
        Frm100.TB2 = vbNullString
        Frm100.TB3 = vbNullString
        Frm100.TB4 = vbNullString
        Frm100.TB5 = vbNullString
        
        MsgBox "Data telah berjaya disimpan.", vbInformation, "Info"
        Frm100.TB1.SetFocus
    End If
End If
End Sub
Private Sub CMD19_Click()
'on error resume next
GM_NEXT_PREV = 1

Call frm100_cash_in_out_header
Call frm100_cash_in_out_report
End Sub
Private Sub CMD2_Click()
'On Error Resume Next
Dim Err(6)
DATA_SAVE = 0

Frm100_LM_EMP_NAME = vbNullString
Frm100_LM_EMP_NO = vbNullString

If Frm100.CB1 = 0 And Frm100.CB2 = 0 Then
    x = x + 1
    Err(x) = "Sila pilih samada [Kemasukan Duit] atau [Pengeluaran Duit]."
End If
If Frm100.CBB1 = vbNullString Then
    x = x + 1
    Err(x) = "Sila pilih [Nama Pekerja]."
End If
If Frm100.TB1 = vbNullString Or (Frm100.TB1 <> vbNullString And Not IsNumeric(Frm100.TB1)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Jumlah]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm100.CB2 = 1 Then

    If Frm100.TB3 = vbNullString Then
        x = x + 1
        Err(x) = "Sila masukkan [Nama]."
    End If
    If Frm100.TB4 = vbNullString Then
        x = x + 1
        Err(x) = "Sila masukkan [No. Kad Pengenalan]."
    End If
    If Frm100.TB5 = vbNullString Then
        x = x + 1
        Err(x) = "Sila masukkan [No. Telefon]."
    End If
    
End If

If x <> 0 Then
    Frm6.Show
    Frm6.Pic1.Cls
    For Y = 1 To x
        Frm6.Pic1.Print Y & " - " & Err(Y)
    Next Y
    Exit Sub
Else
    Note = "Adakah anda ingin simpan data ini ?"
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbNo Then
        Exit Sub
    End If
    If Answer = vbYes Then
    
        If Frm100.CBB1 <> vbNullString Then
            Frm100_LM_EMP_NAME = Split(Frm100.CBB1, "  |  ")(0)
            Frm100_LM_EMP_NO = Split(Frm100.CBB1, "  |  ")(1)
        End If
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 47_account_close where ID='" & Frm100.L5_Text & "'", cn, adOpenKeyset, adLockOptimistic

        If Not rs.EOF Then
            If Frm100.CB1 = 1 Then '0 : Kemasukkan Duit , 1 : Pengeluaran Duit
                rs!jenis = 0
            ElseIf Frm100.CB2 = 1 Then
                rs!jenis = 1
            End If
            rs!tarikh = Frm100.DTPicker1 'Tarikh
            If Frm100.TB1 <> vbNullString Then 'Jumlah
                rs!jumlah = Format(Frm100.TB1, "0.00")
            Else
                rs!jumlah = Format(0, "0.00")
            End If
            If Frm100.TB2 <> vbNullString Then 'Remarks
                rs!remarks = Frm100.TB2
            Else
                rs!remarks = Null
            End If
            If Frm100.CBB1 <> vbNullString Then 'Nama Pekerja
                rs!staff_name = Frm100_LM_EMP_NAME
            Else
                rs!staff_name = Null
            End If
            If Frm100.CBB1 <> vbNullString Then 'Nama Pekerja
                rs!staff_id = Frm100_LM_EMP_NO
            Else
                rs!staff_id = Null
            End If
            If Frm100.TB3 <> vbNullString Then 'Nama
                rs!Nama = UCase(Frm100.TB3)
            Else
                rs!Nama = Null
            End If
            If Frm100.TB4 <> vbNullString Then 'No. kad pengenalan
                rs!no_ic = UCase(Frm100.TB4)
            Else
                rs!no_ic = Null
            End If
            If Frm100.TB5 <> vbNullString Then 'No. telefon
                rs!no_tel = UCase(Frm100.TB5)
            Else
                rs!no_tel = Null
            End If
        
            rs!Status = 1 '0 : Tidak Aktif , 1 : Aktif
            rs!write_timestamp2 = Now
            rs.Update
        End If
        
        rs.Close
        Set rs = Nothing
    
        'User = Split(Frm2.StatusBar1.Panels(3), " : ")(1)
        LogAct_Memory = "[" & Frm100_LM_EMP_NAME & "] Edit data kemasukkan/pengeluaran tunai kedai. ID [" & Frm100.L5_Text & "]."
        LogDate_Memory = DateTime.Date & " " & DateTime.Time$
        Call UpdateLog_Database
        
        GM_NEXT_PREV = 0
        Frm100.L8_Text = -1
        Frm100.L9_Text = 0
        Frm100.L10_Text = 0
        
        Call frm100_cash_in_out_header
        Call frm100_cash_in_out_report
        
        Frm100.Pic1.Visible = False
        Frm100.Pic2.Visible = True
    
        MsgBox "Data telah berjaya disimpan.", vbInformation, "Info"
    End If
End If
End Sub
Private Sub CMD20_Click()
'on error resume next
GM_NEXT_PREV = 0

Call frm100_cash_in_out_header
Call frm100_cash_in_out_report
End Sub
Private Sub CMD3_Click()
'on error resume next
Frm100.Pic1.Visible = False
Frm100.Pic2.Visible = True
End Sub
Private Sub CMD4_Click()
'on error resume next
Note = "Pengeluaran report bagi kemasukkan atau pengeluaran tunai kedai dari " & Frm100.DTPicker2 & " hingga " & Frm100.DTPicker3 & "." & vbCrLf & _
        vbNullString & vbCrLf & _
        "Sistem mungkin mengambil sedikit masa untuk memaparkan semua data." & vbCrLf & _
        vbNullString & vbCrLf & _
        "Teruskan ?"

Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")

If Answer = vbYes Then

    GM_NEXT_PREV = 0
    Frm100.L8_Text = -1
    Frm100.L9_Text = 0
    Frm100.L10_Text = 0
    
    Call frm100_cash_in_out_header
    Call frm100_cash_in_out_report
    
    
End If
End Sub
Private Sub CMD5_Click()
'on error resume next
Frm100.Pic3.Visible = False
End Sub
Private Sub Form_Load()
'on error resume next
Frm100.DTPicker2 = DateTime.Date$
Frm100.DTPicker3 = DateTime.Date$
End Sub
Private Sub Frm100_SM_cetak_voucher_Click()
'on error resume next
DATA_FOUND = 0
Frm100_LM_ID = vbNullString

If Frm100.MSFlexGrid1 <> vbNullString Then

    If IsNumeric(Frm100.MSFlexGrid1) Then
    
        Frm100_LM_ID = Frm100.MSFlexGrid1.TextMatrix(Frm100.MSFlexGrid1, 2) 'No. ID
        
        If Frm100_LM_ID <> vbNullString Then
         
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from 47_account_close where ID='" & Frm100_LM_ID & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                
                If Not IsNull(rs!no_voucher) Then
                
                    G_VOUCHER = rs!no_voucher
                    DATA_FOUND = 1
                    
                Else
                    
                    MsgBox "Tiada voucher dijumpai bagi data ini.", vbInformation, "Info"
                    
                End If

                
            End If
            
            rs.Close
            Set rs = Nothing
        
            If DATA_FOUND = 1 Then
            
                Note = "Cetak voucher pengeluaran duit ini?"
                
                Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
                
                If Answer = vbYes Then
                    
                    Call Frm100_cetak_voucher
                    
                End If
                
            End If
            
        End If
        
    End If

End If
End Sub
Private Sub frm100_sm_edit_data_Click()
'on error resume next
DATA_FOUND = 0
Frm100_LM_STAFF_ID = vbNullString
DATA_PEKERJA_FOUND = 0

If Frm100.MSFlexGrid1 <> vbNullString Then
    Frm100_LM_ID = Frm100.MSFlexGrid1.TextMatrix(Frm100.MSFlexGrid1, 2) 'No. ID
    
    If Frm100_LM_ID <> vbNullString Then
        Note = "Adakah ingin lihat data dan edit data ini ?"
        
        Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
        
        If Answer = vbNo Then
            Exit Sub
        End If
        If Answer = vbYes Then
        
            Call frm100_initial_setting2

            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from 47_account_close where ID='" & Frm100_LM_ID & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                Frm100.L5_Text = Frm100_LM_ID
                If Not IsNull(rs!jenis) Then
                    If rs!jenis = 0 Then
                        Frm100.CB1 = 1
                    ElseIf rs!jenis = 1 Then
                        Frm100.CB2 = 1
                    End If
                End If
                If Not IsNull(rs!tarikh) Then Frm100.DTPicker1 = rs!tarikh 'Tarikh
                If Not IsNull(rs!jumlah) Then Frm100.TB1 = Format(rs!jumlah, "0.00") 'Jumlah
                If Not IsNull(rs!remarks) Then Frm100.TB2 = rs!remarks 'Remarks
                If Not IsNull(rs!staff_id) Then Frm100_LM_STAFF_ID = rs!staff_id
                If Not IsNull(rs!Nama) Then Frm100.TB3 = rs!Nama 'Nama
                If Not IsNull(rs!no_ic) Then Frm100.TB4 = rs!no_ic 'No. kad pengenalan
                If Not IsNull(rs!no_tel) Then Frm100.TB5 = rs!no_tel 'No. telefon
                
                DATA_FOUND = 1
            End If
            
            rs.Close
            Set rs = Nothing
            
            If DATA_FOUND = 1 Then
            
                If Frm100_LM_STAFF_ID <> vbNullString Then
                
                    '### Carian Maklumat Penjual (Data Pekerja) ### - Start
                    
                    Set rs = New ADODB.Recordset
                    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                    rs.Open "select * from employee where NoPekerja='" & Frm100_LM_STAFF_ID & "'", cn, adOpenKeyset, adLockOptimistic
                    
                    If Not rs.EOF Then
                        Frm100_LM_MAKLUMAT_PEKERJA = rs!Samaran & "  |  " & rs!NoPekerja
                        DATA_PEKERJA_FOUND = 1
                    End If
                    
                    rs.Close
                    Set rs = Nothing
                    
                    If DATA_PEKERJA_FOUND = 1 Then
                        On Error GoTo Err_A:
                        Frm100.CBB1 = Frm100_LM_MAKLUMAT_PEKERJA
Restore_A:
                    End If
                    '### Carian Maklumat Penjual (Data Pekerja) ### - End

                End If
                
                Frm100.CBB1.Enabled = True
                Frm100.CBB1.BackColor = &HFFFFFF
            
                Frm100.CMD1.Visible = False
                Frm100.CMD2.Visible = True
                Frm100.CMD3.Visible = True
                
                Frm100.Pic1.Visible = True
                Frm100.Pic2.Visible = False
            End If
            
        End If
    End If
End If
            
Exit Sub
Err_A:
Frm100.CBB1.AddItem Frm100_LM_MAKLUMAT_PEKERJA
Frm100.CBB1 = Frm100_LM_MAKLUMAT_PEKERJA
Resume Restore_A:

End Sub
Private Sub frm100_sm_padam_Click()
'on error resume next
DATA_FOUND = 0

If Frm100.MSFlexGrid1 <> vbNullString Then
    Frm100_LM_ID = Frm100.MSFlexGrid1.TextMatrix(Frm100.MSFlexGrid1, 2) 'No. ID
    
    If Frm100_LM_ID <> vbNullString Then
        Note = "Adakah anda yakin untuk padam data ini ?"
        
        Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
        
        If Answer = vbNo Then
            Exit Sub
        End If
        If Answer = vbYes Then

            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from 47_account_close where ID='" & Frm100_LM_ID & "'", cn, adOpenKeyset, adLockOptimistic
    
            If Not rs.EOF Then

                rs!Status = 0 '0 : Tidak Aktif , 1 : Aktif
                rs!write_timestamp3 = DateTime.Date$
                rs.Update
                
                DATA_FOUND = 1
            End If
                
            rs.Close
            Set rs = Nothing
            
            
            If DATA_FOUND = 1 Then
                
                user = Split(Frm2.StatusBar1.Panels(3), " : ")(1)
                LogAct_Memory = "[" & user & "] Padam data kemasukkan / pengeluaran tunai kedai. ID : " & Frm100_LM_ID & "."
                LogDate_Memory = DateTime.Date & " " & DateTime.Time$
                Call UpdateLog_Database
                
                GM_NEXT_PREV = 0
                Frm100.L8_Text = -1
                Frm100.L9_Text = 0
                Frm100.L10_Text = 0
                
                Call frm100_cash_in_out_header
                Call frm100_cash_in_out_report
                
                MsgBox "Data telah berjaya dipadam.", vbInformation, "Info"
                
            End If
        End If
    End If
End If
End Sub
Private Sub Frm100_SM_terperinci_Click()
'on error resume next
DATA_FOUND = 0

If Frm100.MSFlexGrid1 <> vbNullString Then
    Frm100_LM_ID = Frm100.MSFlexGrid1.TextMatrix(Frm100.MSFlexGrid1, 2) 'No. ID
    
    If Frm100_LM_ID <> vbNullString Then
        Note = "Lihat data terperinci urusan ini ?"
        
        Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
        
        If Answer = vbNo Then
            Exit Sub
        End If
        If Answer = vbYes Then

            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from 47_account_close where ID='" & Frm100_LM_ID & "'", cn, adOpenKeyset, adLockOptimistic
    
            If Not rs.EOF Then
                If Not IsNull(rs!jenis) Then
                    If rs!jenis = 0 Then
                        Frm100_LM_JENIS = "Kemasukkan Duit"
                    ElseIf rs!jenis = 1 Then
                        Frm100_LM_JENIS = "Pengeluaran Duit"
                    End If
                Else
                    Frm100_LM_JENIS = vbnustring
                End If
                If Not IsNull(rs!tarikh) Then
                    Frm100_LM_TARIKH = rs!tarikh
                Else
                    Frm100_LM_TARIKH = vbNullString
                End If
                If Not IsNull(rs!jumlah) Then
                    Frm100_LM_JUMLAH = "RM " & Format(rs!jumlah, "#,##0.00")
                Else
                    Frm100_LM_JUMLAH = "0.00"
                End If
                If Not IsNull(rs!remarks) Then
                    Frm100_LM_REMARKS = rs!remarks
                Else
                    Frm100_LM_REMARKS = vbNullString
                End If
                
                If Not IsNull(rs!staff_name) Then
                    Frm100_LM_STAFF = rs!staff_name
                Else
                    Frm100_LM_STAFF = vbNullString
                End If
                
                If Not IsNull(rs!Nama) Then
                    Frm100_LM_NAMA = rs!Nama
                Else
                    Frm100_LM_NAMA = vbNullString
                End If
                If Not IsNull(rs!no_ic) Then
                    Frm100_LM_IC = rs!no_ic
                Else
                    Frm100_LM_IC = vbNullString
                End If
                If Not IsNull(rs!no_tel) Then
                    Frm100_LM_TEL = rs!no_tel
                Else
                    Frm100_LM_TEL = vbNullString
                End If
                DATA_FOUND = 1
            End If
                
            rs.Close
            Set rs = Nothing
            
            
            If DATA_FOUND = 1 Then
                
                Frm100.L13_Text = vbNullString
                Frm100.L13_Text = "                         Sankyu System                   " & vbCrLf & _
                                "=========================================================" & vbCrLf & _
                                vbNullString & vbCrLf & _
                                vbNullString & vbCrLf & _
                                "Jenis : " & Frm100_LM_JENIS & vbCrLf & _
                                "Tarikh : " & Frm100_LM_TARIKH & vbCrLf & _
                                "Nama Staff Bertugas : " & Frm100_LM_STAFF & vbCrLf & _
                                "Jumlah : " & Frm100_LM_JUMLAH & vbCrLf & _
                                vbNullString & vbCrLf & _
                                "Remarks : " & Frm100_LM_REMARKS & vbCrLf & _
                                vbNullString & vbCrLf & _
                                "Nama : " & Frm100_LM_NAMA & vbCrLf & _
                                "No. Kad Pengenalan : " & Frm100_LM_IC & vbCrLf & _
                                "No. Telefon : " & Frm100_LM_TEL
                
            End If
        End If
    End If
End If
End Sub
Private Sub L3_Text_Click()
'on error resume next
If Frm100.Pic1.Visible = False Then
    Call frm100_initial_setting
    
    Frm100.Pic1.Visible = True
    Frm100.TB1.SetFocus
Else
    Frm100.Pic1.Visible = False
End If
End Sub
Private Sub Label3_Click()
'on error resume next
If Frm100.Pic3.Visible = False Then
    Call frm100_initial_setting
    
    Frm100.Pic3.Visible = True
Else
    Frm100.Pic3.Visible = False
End If
End Sub
Private Sub MSFlexGrid1_DblClick()
'On Error Resume Next
If Frm100.MSFlexGrid1 <> vbNullString Then

    user_level = MDI_frm1.L4_Text
    
    If user_level = "Admin" Or user_level = "HQ" Or user_level = "Developer" Then
    
        Frm100.frm100_sm_edit_data.Enabled = True
        Frm100.frm100_sm_padam.Enabled = True
                
    ElseIf user_level = "Manager" Then
    
        Frm100.frm100_sm_edit_data.Enabled = True
        Frm100.frm100_sm_padam.Enabled = False
        
    Else
    
        Frm100.frm100_sm_edit_data.Enabled = False
        Frm100.frm100_sm_padam.Enabled = False
    
    End If
    
    PopupMenu frm100_pm_menu
Else
    MsgBox "Tiada Data.", vbExclamation, "Info"
End If
End Sub
