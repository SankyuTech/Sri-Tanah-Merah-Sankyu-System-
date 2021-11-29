VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm86 
   Caption         =   "Pengurusan Buku Cek"
   ClientHeight    =   12915
   ClientLeft      =   120
   ClientTop       =   -5175
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
   Icon            =   "Frm86.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   12915
   ScaleWidth      =   23760
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Pic4 
      BorderStyle     =   0  'None
      Height          =   10935
      Left            =   10200
      ScaleHeight     =   10935
      ScaleWidth      =   23535
      TabIndex        =   36
      Top             =   2040
      Visible         =   0   'False
      Width           =   23535
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
         Height          =   10035
         Left            =   120
         TabIndex        =   41
         ToolTipText     =   "Sila ""Double Click"" untuk menu seterusnya."
         Top             =   480
         Width           =   20655
         _ExtentX        =   36433
         _ExtentY        =   17701
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
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Senarai pengeluaran cek."
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
         Height          =   405
         Left            =   240
         TabIndex        =   37
         Top             =   240
         Width           =   7185
      End
   End
   Begin VB.PictureBox Pic3 
      BorderStyle     =   0  'None
      Height          =   9495
      Left            =   6960
      ScaleHeight     =   9495
      ScaleWidth      =   23535
      TabIndex        =   16
      Top             =   2640
      Visible         =   0   'False
      Width           =   23535
      Begin VB.CommandButton CMD5 
         BackColor       =   &H000080FF&
         Caption         =   "Batal"
         Height          =   400
         Left            =   5040
         MaskColor       =   &H00400000&
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   5880
         Visible         =   0   'False
         Width           =   2505
      End
      Begin VB.CommandButton CMD4 
         BackColor       =   &H000080FF&
         Caption         =   "Simpan Data"
         Height          =   400
         Left            =   2160
         MaskColor       =   &H00400000&
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   5880
         Visible         =   0   'False
         Width           =   2505
      End
      Begin VB.CommandButton CMD3 
         BackColor       =   &H000080FF&
         Caption         =   "Simpan Data"
         Height          =   400
         Left            =   3600
         MaskColor       =   &H00400000&
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   5160
         Width           =   2505
      End
      Begin VB.TextBox TB6 
         BackColor       =   &H00FFFFFF&
         Height          =   1440
         Left            =   2025
         MultiLine       =   -1  'True
         TabIndex        =   32
         Text            =   "Frm86.frx":0ECA
         Top             =   2400
         Width           =   7020
      End
      Begin VB.TextBox TB5 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   2025
         TabIndex        =   30
         Text            =   "TB5"
         Top             =   2040
         Width           =   7020
      End
      Begin VB.TextBox TB4 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   2025
         TabIndex        =   28
         Text            =   "TB4"
         Top             =   1680
         Width           =   7020
      End
      Begin VB.TextBox TB3 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   2025
         TabIndex        =   24
         Text            =   "TB3"
         Top             =   1320
         Width           =   7020
      End
      Begin VB.ComboBox CBB2 
         BackColor       =   &H00FFFFFF&
         DataField       =   "Supplier"
         Height          =   360
         ItemData        =   "Frm86.frx":0ECE
         Left            =   2025
         List            =   "Frm86.frx":0ED0
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   4200
         Width           =   7000
      End
      Begin VB.ComboBox CBB1 
         BackColor       =   &H00FFFFFF&
         DataField       =   "Supplier"
         Height          =   360
         ItemData        =   "Frm86.frx":0ED2
         Left            =   2040
         List            =   "Frm86.frx":0ED4
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   480
         Width           =   7000
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   360
         Left            =   2025
         TabIndex        =   22
         Top             =   4560
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
         Format          =   415432704
         CurrentDate     =   41561
      End
      Begin VB.Label L10_Text 
         Caption         =   "L10_Text"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   240
         TabIndex        =   38
         Top             =   5280
         Visible         =   0   'False
         Width           =   1905
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks                 :"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   120
         TabIndex        =   33
         Top             =   2400
         Width           =   1905
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Dibayar Kepada       :"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   120
         TabIndex        =   31
         Top             =   2040
         Width           =   1905
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah *           : RM"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   120
         TabIndex        =   29
         Top             =   1680
         Width           =   1905
      End
      Begin VB.Label L7_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L7_Text"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   2040
         TabIndex        =   27
         Top             =   900
         Width           =   7095
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "No. Akaun *            :"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   120
         TabIndex        =   26
         Top             =   900
         Width           =   2295
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "No. Cek *                :"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   120
         TabIndex        =   25
         Top             =   1320
         Width           =   1905
      End
      Begin VB.Label Label62 
         BackStyle       =   0  'Transparent
         Caption         =   "Tarikh                    :"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   4605
         Width           =   2535
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Pekerja *      :"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   120
         TabIndex        =   21
         Top             =   4215
         Width           =   2295
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Sila masukkan data terperinci bagi pengeluaran cek."
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
         Height          =   405
         Left            =   120
         TabIndex        =   19
         Top             =   120
         Width           =   8055
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Bank  *          :"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   120
         TabIndex        =   18
         Top             =   495
         Width           =   2295
      End
   End
   Begin VB.PictureBox Pic2 
      BorderStyle     =   0  'None
      Height          =   10215
      Left            =   1200
      ScaleHeight     =   10215
      ScaleWidth      =   9975
      TabIndex        =   10
      Top             =   2760
      Visible         =   0   'False
      Width           =   9975
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   9525
         Left            =   240
         TabIndex        =   12
         ToolTipText     =   "Sila ""Double Click"" untuk menu seterusnya."
         Top             =   480
         Width           =   9465
         _ExtentX        =   16695
         _ExtentY        =   16801
         _Version        =   393216
         Rows            =   1
         Cols            =   0
         FixedCols       =   0
         BackColor       =   16777088
         ForeColor       =   0
         BackColorFixed  =   8454016
         BackColorBkg    =   12640511
         GridColor       =   0
         WordWrap        =   -1  'True
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Senarai buku cek."
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
         Height          =   405
         Left            =   360
         TabIndex        =   13
         Top             =   120
         Width           =   7185
      End
   End
   Begin VB.PictureBox Pic1 
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   120
      ScaleHeight     =   1935
      ScaleWidth      =   8175
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   8175
      Begin VB.CommandButton CMD2 
         BackColor       =   &H000080FF&
         Caption         =   "Simpan Data"
         Height          =   400
         Left            =   2760
         MaskColor       =   &H00400000&
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1320
         Width           =   2505
      End
      Begin VB.CommandButton CMD1 
         BackColor       =   &H000080FF&
         Caption         =   "Simpan Data"
         Height          =   400
         Left            =   2760
         MaskColor       =   &H00400000&
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1320
         Width           =   2505
      End
      Begin VB.TextBox TB2 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   1905
         TabIndex        =   5
         Text            =   "TB2"
         Top             =   840
         Width           =   5460
      End
      Begin VB.TextBox TB1 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   1905
         TabIndex        =   3
         Text            =   "TB1"
         Top             =   480
         Width           =   5460
      End
      Begin VB.Label L4_Text 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
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
         TabIndex        =   9
         Top             =   1560
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Pendaftaran buku cek ke dalam sistem."
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
         Height          =   405
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   7185
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "No. Akaun             :"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   120
         TabIndex        =   6
         Top             =   870
         Width           =   1905
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Bank            :"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   120
         TabIndex        =   4
         Top             =   510
         Width           =   1905
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   9495
      Left            =   15960
      ScaleHeight     =   9495
      ScaleWidth      =   23535
      TabIndex        =   0
      Top             =   2640
      Visible         =   0   'False
      Width           =   23535
   End
   Begin VB.Timer Tmr1 
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin VB.Label L8_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Senarai Pengeluaran Cek"
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
      Left            =   8280
      MouseIcon       =   "Frm86.frx":0ED6
      MousePointer    =   99  'Custom
      TabIndex        =   35
      Top             =   0
      Width           =   2295
   End
   Begin VB.Label L6_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pengeluaran Cek"
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
      Left            =   5520
      MouseIcon       =   "Frm86.frx":11E0
      MousePointer    =   99  'Custom
      TabIndex        =   15
      Top             =   0
      Width           =   2295
   End
   Begin VB.Label L5_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Senarai Buku Cek"
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
      MouseIcon       =   "Frm86.frx":14EA
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   0
      Width           =   2295
   End
   Begin VB.Label L3_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pendaftaran Buku Cek"
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
      MouseIcon       =   "Frm86.frx":17F4
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   0
      Width           =   2295
   End
   Begin VB.Menu Frm86_PM_Menu1 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu Frm86_SM_Edit_Data 
         Caption         =   "Edit Data"
      End
      Begin VB.Menu Frm86_SM_Padam_Data 
         Caption         =   "Padam Data"
      End
   End
   Begin VB.Menu Frm86_PM_Menu2 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu Frm86_SM_Edit_Data2 
         Caption         =   "Edit Data Cek"
      End
      Begin VB.Menu Frm86_SM_Padam_Data2 
         Caption         =   "Padam Data"
      End
   End
End
Attribute VB_Name = "Frm86"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CBB1_Change()
'On Error Resume Next
If GLOBAL_DISABLE = 0 Then
    If Frm86.CBB1 <> vbNullString Then
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 25_buku_cek where nama_bank='" & Frm86.CBB1 & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            Frm86.L7_Text = vbNullString
            If Not IsNull(rs!no_akaun) Then Frm86.L7_Text = rs!no_akaun 'No. Akaun
        End If
        
        rs.Close
        Set rs = Nothing
    End If
End If
End Sub
Private Sub CBB1_Click()
'On Error Resume Next
If GLOBAL_DISABLE = 0 Then
    If Frm86.CBB1 <> vbNullString Then
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 25_buku_cek where nama_bank='" & Frm86.CBB1 & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            Frm86.L7_Text = vbNullString
            If Not IsNull(rs!no_akaun) Then Frm86.L7_Text = rs!no_akaun 'No. Akaun
        End If
        
        rs.Close
        Set rs = Nothing
    End If
End If
End Sub
Private Sub CMD1_Click()
'On Error Resume Next
Dim Err(5)

If Frm86.TB1 = vbNullString Then
    x = x + 1
    Err(x) = "Sila Masukkan [Nama Bank]."
End If
If Frm86.TB2 = vbNullString Then
    x = x + 1
    Err(x) = "Sila Masukkan [No. Akaun]."
End If

If x <> 0 Then
    Frm6.Show
    Frm6.Pic1.Cls
    For Y = 1 To x
        Frm6.Pic1.Print Y & " - " & Err(Y)
    Next Y
    Exit Sub
Else
    Note = "Adakah Anda Ingin Masukkan Data Ini ?"
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbNo Then
        Exit Sub
    End If
    If Answer = vbYes Then
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 25_buku_cek where nama_bank='" & UCase(Frm86.TB1) & "'", cn, adOpenKeyset, adLockOptimistic
        
        If rs.EOF Then
            rs.AddNew
            rs!nama_bank = UCase(Frm86.TB1) 'Nama Bank Bagi Buku Cek Ini
            rs!no_akaun = UCase(Frm86.TB2) 'No. Akaun Bagi No. Bank Ini
            rs.Update
        Else
            MsgBox "Nama Bank [" & UCase(Frm86.TB1) & "] Sudah Didaftarkan Sebelum Ini.", vbInformation, "Info"
        End If
        
        rs.Close
        Set rs = Nothing
        
        user = MDI_frm1.L3_Text
        LogAct_Memory = "[" & user & "] Pendaftaran Buku Cek , Nama Bank [" & UCase(Frm86.TB1) & "]."
        LogDate_Memory = DateTime.Date & " " & DateTime.Time$
        Call UpdateLog_Database
        
        Call Frm86_Initial_Setting
        MsgBox "Data Telah Berjaya Disimpan.", vbInformation, "Info"
    End If
End If
End Sub
Private Sub CMD2_Click()
'On Error Resume Next
Dim Err(5)
DATA_SAVE = 0
DATA_OK = 0

If Frm86.TB1 = vbNullString Then
    x = x + 1
    Err(x) = "Sila Masukkan [Nama Bank]."
End If
If Frm86.TB2 = vbNullString Then
    x = x + 1
    Err(x) = "Sila Masukkan [No. Akaun]."
End If

If x <> 0 Then
    Frm6.Show
    Frm6.Pic1.Cls
    For Y = 1 To x
        Frm6.Pic1.Print Y & " - " & Err(Y)
    Next Y
    Exit Sub
Else
    Note = "Adakah Anda Ingin Masukkan Data Ini ?"
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbNo Then
        Exit Sub
    End If
    If Answer = vbYes Then
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 25_buku_cek where nama_bank='" & UCase(Frm86.TB1) & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Frm86.L4_Text <> rs!ID Then
                DATA_OK = 1
                MsgBox "Nama Bank [" & UCase(Frm86.TB1) & "] Sudah Didaftarkan Sebelum Ini.", vbInformation, "Info"
            End If
        End If
        
        rs.Close
        Set rs = Nothing
        
        If DATA_OK = 0 Then
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from 25_buku_cek where ID='" & Frm86.L4_Text & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                rs!nama_bank = UCase(Frm86.TB1) 'Nama Bank Bagi Buku Cek Ini
                rs!no_akaun = UCase(Frm86.TB2) 'No. Akaun Bagi No. Bank Ini
                DATA_SAVE = 1
                rs.Update
            End If
            
            rs.Close
            Set rs = Nothing
            
            If DATA_SAVE = 1 Then
                user = MDI_frm1.L3_Text
                LogAct_Memory = "[" & user & "] Edit Data Buku Cek , Nama Bank [" & UCase(Frm86.TB1) & "]."
                LogDate_Memory = DateTime.Date & " " & DateTime.Time$
                Call UpdateLog_Database
                
                Call Frm86_Senarai_Buku_Cek_Header
                Call Frm86_Senarai_Buku_Cek
                
                Frm86.Pic2.Visible = True
                Frm86.Pic1.Visible = False
                
                MsgBox "Data Telah Berjaya Disimpan.", vbInformation, "Info"
            End If
        End If
    End If
End If
End Sub
Private Sub CMD3_Click()
'On Error Resume Next
Dim Err(5)
DATA_SAVE = 0
DATA_OK = 0

If Frm86.CBB1 = vbNullString Then
    x = x + 1
    Err(x) = "Sila Pilih [Nama Bank]."
End If
If Frm86.L7_Text = vbNullString Then
    x = x + 1
    Err(x) = "Tiada Maklumat [No. Akaun]."
End If
If Frm86.TB3 = vbNullString Then
    x = x + 1
    Err(x) = "Sila Isi [No. Cek]."
End If
If Frm86.TB4 = vbNullString Or (Frm86.TB4 <> vbNullString And Not IsNumeric(Frm86.TB4)) Then
    x = x + 1
    Err(x) = "Sila Masukkan [Jumlah]. Hanya NOMBOR Dibenarkan Dalam Ruangan Ini."
End If
If Frm86.CBB2 = vbNullString Then
    x = x + 1
    Err(x) = "Sila Pilih [Nama Pekerja]."
End If

If x <> 0 Then
    Frm6.Show
    Frm6.Pic1.Cls
    For Y = 1 To x
        Frm6.Pic1.Print Y & " - " & Err(Y)
    Next Y
    Exit Sub
Else
    Note = "Adakah Anda Ingin Masukkan Data Ini ?"
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbNo Then
        Exit Sub
    End If
    If Answer = vbYes Then
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 26_senarai_cek where nama_bank='" & Frm86.CBB1 & "' AND no_cek='" & UCase(Frm86.TB3) & "'", cn, adOpenKeyset, adLockOptimistic
        
        If rs.EOF Then
            rs.AddNew
            If Frm86.CBB1 <> vbNullString Then
                rs!nama_bank = Frm86.CBB1 'Nama Bank
            Else
                rs!nama_bank = Null 'Nama Bank
            End If
            If Frm86.L7_Text <> vbNullString Then
                rs!no_akaun = Frm86.L7_Text 'No. Akaun Bank
            Else
                rs!no_akaun = Null 'No. Akaun Bank
            End If
            If Frm86.TB3 <> vbNullString Then
                rs!no_cek = UCase(Frm86.TB3) 'No. Akaun Bank
            Else
                rs!no_cek = Null 'No. Akaun Bank
            End If
            If Frm86.DTPicker1 <> vbNullString Then
                rs!tarikh = Frm86.DTPicker1 'Tarikh
            Else
                rs!tarikh = Null 'Tarikh
            End If
            If Frm86.TB4 <> vbNullString Then
                rs!jumlah = Format(Frm86.TB4, "0.00") 'Jumlah (RM)
            Else
                rs!jumlah = Null 'Jumlah (RM)
            End If
            If Frm86.TB5 <> vbNullString Then
                rs!penerima = Frm86.TB5 'Nama Penerima
            Else
                rs!penerima = Null 'Nama Penerima
            End If
            If Frm86.TB6 <> vbNullString Then
                rs!remarks = Frm86.TB6 'Remarks
            Else
                rs!remarks = Null 'Remarks
            End If
            If Frm86.CBB2 <> vbNullString Then
                Frm86_LM_EMP_NO = Split(Frm86.CBB2, "  |  ")(1)
                Frm86_LM_EMP_NAME = Split(Frm86.CBB2, "  |  ")(0)
                rs!no_pekerja = Frm86_LM_EMP_NO 'No. Pekerja
            End If
            rs.Update
            DATA_SAVE = 1
        Else
            MsgBox "No Cek [" & UCase(Frm86.TB3) & "] Bagi Bank [" & Frm86.CBB1 & "] Telah Disimpan Sebelum Ini.", vbInformation, "Info"
        End If
        
        rs.Close
        Set rs = Nothing
        
        If DATA_SAVE = 1 Then
            'User = MDI_frm1.L3_Text
            LogAct_Memory = "[" & Frm86_LM_EMP_NAME & "] Pengeluaran Cek , Nama Bank [" & UCase(Frm86.CBB1) & "], No Cek [" & UCase(Frm86.TB3) & "]"
            LogDate_Memory = DateTime.Date & " " & DateTime.Time$
            Call UpdateLog_Database
            
            Call Frm86_Initial_Setting
            MsgBox "Data Telah Berjaya Disimpan.", vbInformation, "Info"
        End If
    End If
End If
End Sub
Private Sub CMD4_Click()
'On Error Resume Next
Dim Err(5)
DATA_SAVE = 0
DATA_OK = 0

If Frm86.CBB1 = vbNullString Then
    x = x + 1
    Err(x) = "Sila Pilih [Nama Bank]."
End If
If Frm86.L7_Text = vbNullString Then
    x = x + 1
    Err(x) = "Tiada Maklumat [No. Akaun]."
End If
If Frm86.TB3 = vbNullString Then
    x = x + 1
    Err(x) = "Sila Isi [No. Cek]."
End If
If Frm86.TB4 = vbNullString Or (Frm86.TB4 <> vbNullString And Not IsNumeric(Frm86.TB4)) Then
    x = x + 1
    Err(x) = "Sila Masukkan [Jumlah]. Hanya NOMBOR Dibenarkan Dalam Ruangan Ini."
End If
If Frm86.CBB2 = vbNullString Then
    x = x + 1
    Err(x) = "Sila Pilih [Nama Pekerja]."
End If

If x <> 0 Then
    Frm6.Show
    Frm6.Pic1.Cls
    For Y = 1 To x
        Frm6.Pic1.Print Y & " - " & Err(Y)
    Next Y
    Exit Sub
Else
    Note = "Adakah Anda Ingin Masukkan Data Ini ?"
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbNo Then
        Exit Sub
    End If
    If Answer = vbYes Then

        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 26_senarai_cek where nama_bank='" & Frm86.CBB1 & "' AND no_cek='" & UCase(Frm86.TB3) & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Frm86.L10_Text <> rs!ID Then
                MsgBox "No Cek [" & UCase(Frm86.TB3) & "] Bagi Bank [" & Frm86.CBB1 & "] Telah Disimpan Sebelum Ini.", vbInformation, "Info"
                DATA_OK = 1
            End If
        End If
        
        rs.Close
        Set rs = Nothing
        
        If DATA_OK = 0 Then
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from 26_senarai_cek where id='" & Frm86.L10_Text & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                If Frm86.CBB1 <> vbNullString Then
                    rs!nama_bank = Frm86.CBB1 'Nama Bank
                Else
                    rs!nama_bank = Null 'Nama Bank
                End If
                If Frm86.L7_Text <> vbNullString Then
                    rs!no_akaun = Frm86.L7_Text 'No. Akaun Bank
                Else
                    rs!no_akaun = Null 'No. Akaun Bank
                End If
                If Frm86.TB3 <> vbNullString Then
                    rs!no_cek = UCase(Frm86.TB3) 'No. Akaun Bank
                Else
                    rs!no_cek = Null 'No. Akaun Bank
                End If
                If Frm86.DTPicker1 <> vbNullString Then
                    rs!tarikh = Frm86.DTPicker1 'Tarikh
                Else
                    rs!tarikh = Null 'Tarikh
                End If
                If Frm86.TB4 <> vbNullString Then
                    rs!jumlah = Format(Frm86.TB4, "0.00") 'Jumlah (RM)
                Else
                    rs!jumlah = Null 'Jumlah (RM)
                End If
                If Frm86.TB5 <> vbNullString Then
                    rs!penerima = Frm86.TB5 'Nama Penerima
                Else
                    rs!penerima = Null 'Nama Penerima
                End If
                If Frm86.TB6 <> vbNullString Then
                    rs!remarks = Frm86.TB6 'Remarks
                Else
                    rs!remarks = Null 'Remarks
                End If
                If Frm86.CBB2 <> vbNullString Then
                    Frm86_LM_EMP_NO = Split(Frm86.CBB2, "  |  ")(1)
                    Frm86_LM_EMP_NAME = Split(Frm86.CBB2, "  |  ")(0)
                    rs!no_pekerja = Frm86_LM_EMP_NO 'No. Pekerja
                End If
                rs.Update
                DATA_SAVE = 1
            End If
            
            rs.Close
            Set rs = Nothing
        End If
        
        If DATA_SAVE = 1 Then
            'User = MDI_frm1.L3_Text
            LogAct_Memory = "[" & Frm86_LM_EMP_NAME & "] Edit Data Cek , Nama Bank [" & UCase(Frm86.CBB1) & "], No Cek [" & UCase(Frm86.TB3) & "]"
            LogDate_Memory = DateTime.Date & " " & DateTime.Time$
            Call UpdateLog_Database
            
            Call Frm86_Initial_Setting
            MsgBox "Data Telah Berjaya Disimpan.", vbInformation, "Info"
        End If
    End If
End If
End Sub
Private Sub CMD5_Click()
'on error resume next
Note = "Adakah Anda Ingin Batalkan Urusan Edit Ini ?"

Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")

If Answer = vbNo Then
    Exit Sub
End If
If Answer = vbYes Then
    Frm86.Pic3.Visible = False
End If
End Sub
Private Sub Frm86_SM_Edit_Data_Click()
'on error resume next
DATA_FOUND = 0

If Frm86.MSFlexGrid1 <> vbNullString Then
    Frm86_LM_ID = Frm86.MSFlexGrid1.TextMatrix(Frm86.MSFlexGrid1, 2) 'No. ID
    
    If Frm86_LM_ID <> vbNullString Then
        Call Frm86_Initial_Setting
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 25_buku_cek where ID='" & Frm86_LM_ID & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Not IsNull(rs!ID) Then Frm86.L4_Text = rs!ID
            If Not IsNull(rs!nama_bank) Then Frm86.TB1 = rs!nama_bank 'Nama Bank Bagi Buku Cek Ini
            If Not IsNull(rs!no_akaun) Then Frm86.TB2 = rs!no_akaun 'No. Akaun Bagi No. Bank Ini
            DATA_FOUND = 1
        End If
        
        rs.Close
        Set rs = Nothing
        
        If DATA_FOUND = 1 Then
            Frm86.CMD1.Visible = False
            Frm86.CMD2.Visible = True
            
            Frm86.Pic1.Visible = True
            Frm86.Pic2.Visible = False
        End If
    End If
End If
End Sub
Private Sub Frm86_SM_Edit_Data2_Click()
'on error resume next
DATA_FOUND = 0
DATA_PEKERJA_FOUND = 0
Frm86_LM_No_PEKERJA = vbNullString
Frm86_LM_NAMA_BANK = vbNullString

If Frm86.MSFlexGrid2 <> vbNullString Then
    Frm86_LM_ID = Frm86.MSFlexGrid2.TextMatrix(Frm86.MSFlexGrid2, 2) 'No. ID
    
    If Frm86_LM_ID <> vbNullString Then
        Call Frm86_Initial_Setting
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 26_senarai_cek where ID='" & Frm86_LM_ID & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            GLOBAL_DISABLE = 1
            Frm86.L10_Text = Frm86_LM_ID
            If Not IsNull(rs!no_akaun) Then Frm86.L7_Text = rs!no_akaun 'No. Akaun Bank
            If Not IsNull(rs!no_cek) Then Frm86.TB3 = rs!no_cek 'No. Cek
            If Not IsNull(rs!tarikh) Then Frm86.DTPicker1 = rs!tarikh 'Tarikh
            If Not IsNull(rs!jumlah) Then Frm86.TB4 = rs!jumlah 'Jumlah (RM)
            If Not IsNull(rs!penerima) Then Frm86.TB5 = rs!penerima 'Nama Penerima
            If Not IsNull(rs!remarks) Then Frm86.TB6 = rs!remarks 'Remarks
            If Not IsNull(rs!no_pekerja) Then
                Frm86_LM_No_PEKERJA = rs!no_pekerja  'No. Pekerja
            End If
            If Not IsNull(rs!nama_bank) Then
                Frm86_LM_NAMA_BANK = rs!nama_bank  'Nama Bank
            End If
            GLOBAL_DISABLE = 0
            DATA_FOUND = 1
        End If
        
        rs.Close
        Set rs = Nothing
        
        If DATA_FOUND = 1 Then
            '### Carian Maklumat Penjual (Data Pekerja) ### - Start
            If Frm86_LM_No_PEKERJA <> vbNullString Then
                DATA_PEKERJA_FOUND = 0
                
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from employee where NoPekerja='" & Frm86_LM_No_PEKERJA & "'", cn, adOpenKeyset, adLockOptimistic
                
                If Not rs.EOF Then
                    Frm86_LM_MAKLUMAT_PEKERJA = rs!Samaran & "  |  " & rs!NoPekerja
                    DATA_PEKERJA_FOUND = 1
                End If
                
                rs.Close
                Set rs = Nothing
                
                If DATA_PEKERJA_FOUND = 1 Then
                    On Error GoTo Err_A:
                    Frm86.CBB2 = Frm86_LM_MAKLUMAT_PEKERJA
Restore_A:
                End If
            End If
            '### Carian Maklumat Penjual (Data Pekerja) ### - End
            
            If Frm86_LM_NAMA_BANK <> vbNullString Then
                On Error GoTo Err_B:
                Frm86.CBB1 = Frm86_LM_NAMA_BANK
Restore_B:
            End If
            
            Frm86.CMD3.Visible = False
            Frm86.CMD4.Visible = True
            Frm86.CMD5.Visible = True

            Frm86.Pic3.Visible = True
            Frm86.Pic4.Visible = False
        End If
    End If
End If

Exit Sub
Err_A:
Frm86.CBB2.AddItem Frm86_LM_MAKLUMAT_PEKERJA
Frm86.CBB2 = Frm86_LM_MAKLUMAT_PEKERJA
Resume Restore_A:

Exit Sub
Err_B:
Frm86.CBB1.AddItem Frm86_LM_NAMA_BANK
Frm86.CBB1 = Frm86_LM_NAMA_BANK
Resume Restore_B:
End Sub
Private Sub Frm86_SM_Padam_Data2_Click()
'on error resume next
DATA_FOUND = 0

If Frm86.MSFlexGrid2 <> vbNullString Then
    Frm86_LM_ID = Frm86.MSFlexGrid2.TextMatrix(Frm86.MSFlexGrid2, 2) 'No. ID
    
    If Frm86_LM_ID <> vbNullString Then
        Note = "Adakah Anda Ingin Padamkan Data Ini ?"
        
        Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
        
        If Answer = vbNo Then
            Exit Sub
        End If
        If Answer = vbYes Then
    
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from 26_senarai_cek where ID='" & Frm86_LM_ID & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                If Not IsNull(rs!nama_bank) Then Frm86_LM_NAMA_BANK = rs!nama_bank
                If Not IsNull(rs!no_cek) Then Frm86_LM_No_CEK = rs!no_cek
                rs.Delete
                rs.Update
                DATA_FOUND = 1
            End If
            
            rs.Close
            Set rs = Nothing
            
            If DATA_FOUND = 1 Then
                user = MDI_frm1.L3_Text
                LogAct_Memory = "[" & user & "] Padam Data Cek , Nama Bank [" & Frm86_LM_NAMA_BANK & "], No Cek [" & Frm86_LM_No_CEK & "]"
                LogDate_Memory = DateTime.Date & " " & DateTime.Time$
                Call UpdateLog_Database
                
                Note = "Data Telah Berjaya Disimpan" & vbCrLf & _
                        "Refresh Senarai Pengeluaran Cek ?"
                
                Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
                
                If Answer = vbNo Then
                    Exit Sub
                End If
                If Answer = vbYes Then
                    Call Frm86_Senarai_Cek_Header
                    Call Frm86_Senarai_Cek
                End If
            End If
            
        End If
    End If
End If
End Sub
Private Sub L3_Text_Click()
'on error resume next
If Frm86.Pic1.Visible = False Then
    Call Frm86_Initial_Setting
    
    Frm86.Pic1.Visible = True
Else
    Frm86.Pic1.Visible = False
End If
End Sub
Private Sub L5_Text_Click()
'on error resume next
If Frm86.Pic2.Visible = False Then
    Call Frm86_Initial_Setting
    Call Frm86_Senarai_Buku_Cek_Header
    Call Frm86_Senarai_Buku_Cek
    
    Frm86.Pic2.Visible = True
Else
    Frm86.Pic2.Visible = False
End If
End Sub
Private Sub L6_Text_Click()
'on error resume next
If Frm86.Pic3.Visible = False Then
    Call Frm86_Initial_Setting
    
    Frm86.Pic3.Visible = True
Else
    Frm86.Pic3.Visible = False
End If
End Sub
Private Sub L8_Text_Click()
'on error resume next
If Frm86.Pic4.Visible = False Then
    Note = "Sistem Akan Mengambil Sedikit Masa Untuk Mengeluarkan Senarai Ini." & vbCrLf & _
            "Teruskan ?"
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbNo Then
        Exit Sub
    End If
    If Answer = vbYes Then
        Call Frm86_Initial_Setting
        Call Frm86_Senarai_Cek_Header
        Call Frm86_Senarai_Cek
        
        Frm86.Pic4.Visible = True
    End If
Else
    Frm86.Pic4.Visible = False
End If
End Sub

Private Sub MSFlexGrid1_DblClick()
'on error resume next
If Frm86.MSFlexGrid1 <> vbNullString Then
    PopupMenu Frm86_PM_Menu1
Else
    MsgBox "Tiada Data.", vbExclamation, "Info"
End If
End Sub
Private Sub MSFlexGrid2_DblClick()
'on error resume next
If Frm86.MSFlexGrid2 <> vbNullString Then
    PopupMenu Frm86_PM_Menu2
Else
    MsgBox "Tiada Data.", vbExclamation, "Info"
End If
End Sub
