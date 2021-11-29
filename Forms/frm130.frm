VERSION 5.00
Begin VB.Form frm130 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Bayaran"
   ClientHeight    =   7140
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9015
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
   LinkTopic       =   "Form1"
   ScaleHeight     =   7140
   ScaleWidth      =   9015
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CMD2 
      BackColor       =   &H80000000&
      Caption         =   "KELUAR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4320
      Picture         =   "frm130.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   6000
      Width           =   2055
   End
   Begin VB.CommandButton CMD1 
      BackColor       =   &H80000000&
      Caption         =   "SIMPAN"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2040
      Picture         =   "frm130.frx":25CA
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   6000
      Width           =   2055
   End
   Begin VB.TextBox TB33 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   900
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   15
      Text            =   "0.00"
      Top             =   360
      Width           =   5205
   End
   Begin VB.TextBox TB27 
      BackColor       =   &H00FFFFFF&
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
      Left            =   2160
      TabIndex        =   5
      Text            =   "0.00"
      Top             =   2160
      Width           =   1380
   End
   Begin VB.TextBox TB28 
      BackColor       =   &H00FFFFFF&
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
      Left            =   2160
      TabIndex        =   4
      Text            =   "0.00"
      Top             =   2520
      Width           =   1380
   End
   Begin VB.TextBox TB29 
      BackColor       =   &H00FFFFFF&
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
      Left            =   4680
      TabIndex        =   3
      Text            =   "0.00"
      Top             =   3075
      Width           =   1500
   End
   Begin VB.TextBox TB32 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "0.00"
      Top             =   5400
      Width           =   1740
   End
   Begin VB.TextBox TB21 
      BackColor       =   &H00FFFFFF&
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
      Left            =   6240
      TabIndex        =   1
      Text            =   "0.00"
      Top             =   2160
      Width           =   1380
   End
   Begin VB.ComboBox CBB2 
      BackColor       =   &H00FFFFFF&
      DataField       =   "Supplier"
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
      ItemData        =   "frm130.frx":4B94
      Left            =   4680
      List            =   "frm130.frx":4B96
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   3480
      Width           =   3885
   End
   Begin VB.Label L8_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L8_Text"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1560
      TabIndex        =   28
      Top             =   6360
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Label L41_Text 
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   480
      TabIndex        =   27
      Top             =   6360
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah potongan kad kredit/debit : RM"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   240
      TabIndex        =   24
      Top             =   4920
      Width           =   4395
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cukai GST caj perkhidmatan kad kredit/debit : RM"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   240
      TabIndex        =   23
      Top             =   4560
      Width           =   4395
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah caj perkidmatan kad kredit/debit : RM"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   240
      TabIndex        =   22
      Top             =   4200
      Width           =   4395
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Caj perkhidmatan kad kredit/debit : %"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   240
      TabIndex        =   21
      Top             =   3840
      Width           =   4395
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Jenis kad kredit/debit :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   240
      TabIndex        =   20
      Top             =   3480
      Width           =   4395
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah bayaran dai kad kredit / debit : RM"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   240
      TabIndex        =   19
      Top             =   3120
      Width           =   4395
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Simpanan di kedai : RM"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   3720
      TabIndex        =   18
      Top             =   2205
      Width           =   2475
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Online Transfer : RM"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   240
      TabIndex        =   17
      Top             =   2565
      Width           =   1875
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah  RM :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   765
      Left            =   240
      TabIndex        =   16
      Top             =   480
      Width           =   3315
   End
   Begin VB.Shape Shape10 
      BorderWidth     =   2
      Height          =   4455
      Left            =   120
      Top             =   1440
      Width           =   8655
   End
   Begin VB.Label Label70 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tunai : RM"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   600
      TabIndex        =   14
      Top             =   2205
      Width           =   1515
   End
   Begin VB.Label Label76 
      BackStyle       =   0  'Transparent
      Caption         =   "Maklumat Cara Bayaran"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   360
      TabIndex        =   13
      Top             =   1560
      Width           =   3585
   End
   Begin VB.Label Label81 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah Bayaran    RM :"
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
      Left            =   4080
      TabIndex        =   12
      Top             =   5430
      Width           =   2715
   End
   Begin VB.Label L31_Text 
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   4680
      TabIndex        =   11
      Top             =   3840
      Width           =   2040
   End
   Begin VB.Label L32_Text 
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   4680
      TabIndex        =   10
      Top             =   4200
      Width           =   2040
   End
   Begin VB.Label Label82 
      BackStyle       =   0  'Transparent
      Caption         =   "Simpanan Duit Di Kedai Sebanyak : RM"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   3480
      TabIndex        =   9
      Top             =   1800
      Width           =   3435
   End
   Begin VB.Label L26_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   6840
      TabIndex        =   8
      Top             =   1800
      Width           =   1635
   End
   Begin VB.Label L82_Text 
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   4680
      TabIndex        =   7
      Top             =   4920
      Width           =   2040
   End
   Begin VB.Label L81_Text 
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   4680
      TabIndex        =   6
      Top             =   4560
      Width           =   2040
   End
   Begin VB.Shape Shape3 
      Height          =   2175
      Left            =   240
      Top             =   3000
      Width           =   8415
   End
End
Attribute VB_Name = "frm130"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CBB2_Change()
'on error resume next
If frm130.CBB2 <> vbNullString Then
    If frm130.L41_Text = "0" Then
    'If GLOBAL_DISABLE <> 1 Then
    
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 74_cas_kad_kredit where jenis_kad='" & frm130.CBB2 & "' AND status='" & 1 & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Not IsNull(rs!cas_kad) Then
                If IsNumeric(rs!cas_kad) Then
                    frm130.L31_Text = rs!cas_kad
                Else
                    frm130.L31_Text = "0.00"
                End If
            Else
                frm130.L31_Text = "0.00"
            End If
            
        End If
        
        rs.Close
        Set rs = Nothing
    
    End If
End If
End Sub

Private Sub CBB2_Click()
'on error resume next
If frm130.CBB2 <> vbNullString Then
    
    If frm130.L41_Text = "0" Then
    'If GLOBAL_DISABLE <> 1 Then
    
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 74_cas_kad_kredit where jenis_kad='" & frm130.CBB2 & "' AND status='" & 1 & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Not IsNull(rs!cas_kad) Then
                If IsNumeric(rs!cas_kad) Then
                    frm130.L31_Text = rs!cas_kad
                Else
                    frm130.L31_Text = "0.00"
                End If
            Else
                frm130.L31_Text = "0.00"
            End If
            
        End If
        
        rs.Close
        Set rs = Nothing
    
    End If
End If
End Sub
Private Sub CMD1_Click()
'on error resume next
Dim Err(30)

Dim frm130_LM_JUMLAH_BAYARAN As Double
Dim frm130_LM_HARGA As Double
Dim frm130_LM_JUMLAH_SIMPANAN As Double
Dim frm130_LM_GUNA_SIMPAN As Double

If frm130.TB27 = vbNullString Or (frm130.TB27 <> vbNullString And Not IsNumeric(frm130.TB27)) Then
    x = x + 1
    Err(x) = "Hanya NOMBOR dibenarkan dalam ruangan bayaran secara TUNAI. Sila masukkan 0 jika tiada bayaran secara tunai."
End If
If frm130.TB28 = vbNullString Or (frm130.TB28 <> vbNullString And Not IsNumeric(frm130.TB28)) Then
    x = x + 1
    Err(x) = "Hanya NOMBOR dibenarkan dalam ruangan bayaran secara ONLINE TRANSFER. Sila masukkan 0 jka tiada bayaran secara online traansfer."
End If
If frm130.TB29 = vbNullString Or (frm130.TB29 <> vbNullString And Not IsNumeric(frm130.TB29)) Then
    x = x + 1
    Err(x) = "Hanya NOMBOR dibenarkan dalam ruangan bayaran secara KAD KREDIT. Sila masukkan 0 jika tiada bayaran secara kad kredit."
End If
If frm130.TB21 = vbNullString Or (frm130.TB21 <> vbNullString And Not IsNumeric(frm130.TB21)) Then
    x = x + 1
    Err(x) = "Hanya NOMBOR dibenarkan dalam ruangan bayaran secara Duit Simpanan Di Kedai. Sila masukkan 0 jika tiada bayaran secara simpanan di kedai."
End If

'Error bagi penggunaan kad kredit - Start
If frm130.TB29 <> "0.00" And IsNumeric(frm130.TB29) Then

    If frm130.CBB2 = vbNullString Then
        x = x + 1
        Err(x) = "Sila pilih jenis kad kredit/debit"
    End If
    If frm130.L31_Text = vbNullString Or (frm130.L31_Text <> vbNullString And Not IsNumeric(frm130.L31_Text)) Then
        x = x + 1
        Err(x) = "Tiada maklumat bagi caj perkhidmatan kad kredit/debit."
    End If
    If frm130.L32_Text = vbNullString Or (frm130.L32_Text <> vbNullString And Not IsNumeric(frm130.L32_Text)) Then
        x = x + 1
        Err(x) = "Tiada maklumat bagi jumlah caj perkhidmatan kad kredit/debit."
    End If
    If frm130.L81_Text = vbNullString Or (frm130.L81_Text <> vbNullString And Not IsNumeric(frm130.L81_Text)) Then
        x = x + 1
        Err(x) = "Tiada maklumat bagi jumlah cukai GST bagi caj perkhidmatan kad kredit/debit."
    End If
    If frm130.L81_Text = vbNullString Or (frm130.L81_Text <> vbNullString And Not IsNumeric(frm130.L81_Text)) Then
        x = x + 1
        Err(x) = "Tiada maklumat bagi jumlah potongan kad kredit/debit."
    End If
    
End If
'Error bagi penggunaan kad kredit - End


If (frm130.TB32 <> vbNullString And IsNumeric(frm130.TB32)) And (frm130.TB33 <> vbNullString And IsNumeric(frm130.TB33)) Then
    frm130_LM_JUMLAH_BAYARAN = frm130.TB32 'Jumlah Bayaran
    frm130_LM_HARGA = frm130.TB33 'Harga Keseluruhan
    
    If frm130_LM_JUMLAH_BAYARAN <> frm130_LM_HARGA Then
        x = x + 1
        Err(x) = "Jumlah bayaran tidak sama dengan jumlah harga barang."
    End If
End If

If (frm130.TB21 <> vbNullString And IsNumeric(frm130.TB21)) And (frm130.L26_Text <> vbNullString And IsNumeric(frm130.L26_Text)) Then
    frm130_LM_JUMLAH_SIMPANAN = frm130.L26_Text  'Jumlah Simpanan Yang Ada
    frm130_LM_GUNA_SIMPAN = frm130.TB21  'Jumlah Simpanan Yang Hendak Digunakan
    
    If frm130_LM_GUNA_SIMPAN > frm130_LM_JUMLAH_SIMPANAN Then
        x = x + 1
        Err(x) = "Jumlah simpanan yang hendak digunakan melebihi simpanan terkumpul yang ada."
    End If
End If

If x <> 0 Then
    Frm6.Show vbModal
    Frm6.Pic1.Cls
    For Y = 1 To x
        Frm6.Pic1.Print Y & " - " & Err(Y)
    Next Y
    Exit Sub
Else
    frm130.Hide
End If
End Sub

Private Sub CMD2_Click()
'on error resume next
frm130.Hide
End Sub

Private Sub Form_Load()
'On Error Resume Next
'G_RATE_GST = 6

frm130.CBB2.Clear

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 74_cas_kad_kredit where status = 1 order by jenis_kad ASC", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    If Not IsNull(rs!jenis_kad) Then frm130.CBB2.AddItem rs!jenis_kad
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing
End Sub

Private Sub L31_Text_Change()
'On Error Resume Next
Call Frm130_kira_caj_kad_kredit
End Sub
Private Sub L32_Text_Change()
'On Error Resume Next
Call Frm130_kira_caj_gst_kad_kredit
Call Frm130_kira_potongan_kad_kredit
End Sub

Private Sub L81_Text_Change()
'On Error Resume Next
Call Frm130_kira_potongan_kad_kredit
End Sub

Private Sub TB21_Change()
'On Error Resume Next
Call frm130_kira_jumlah_bayaran
End Sub
Private Sub TB27_Change()
'On Error Resume Next
Call frm130_kira_jumlah_bayaran
End Sub
Private Sub TB28_Change()
'On Error Resume Next
Call frm130_kira_jumlah_bayaran
End Sub
Private Sub TB29_Change()
'On Error Resume Next
Call frm130_kira_jumlah_bayaran
Call Frm130_kira_caj_kad_kredit
Call Frm130_kira_potongan_kad_kredit
End Sub

Private Sub TB33_Change()
'On Error Resume Next
frm130.TB27 = Format(frm130.TB33, "#,##0.00")
End Sub
