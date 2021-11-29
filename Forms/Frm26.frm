VERSION 5.00
Begin VB.Form Frm26 
   BackColor       =   &H80000003&
   Caption         =   "Maklumat Pembeli / Penjual"
   ClientHeight    =   4500
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9000
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
   Icon            =   "Frm26.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4500
   ScaleWidth      =   9000
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CMD2 
      BackColor       =   &H80000004&
      Caption         =   "Padam Data"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1245
      Left            =   7320
      MaskColor       =   &H00400000&
      Picture         =   "Frm26.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Padam Data Yang Telah Diisi"
      Top             =   240
      Width           =   1545
   End
   Begin VB.TextBox TB2 
      Height          =   360
      Left            =   1440
      TabIndex        =   1
      Text            =   "TB2"
      Top             =   840
      Width           =   5655
   End
   Begin VB.TextBox TB1 
      Height          =   360
      Left            =   1440
      TabIndex        =   0
      Text            =   "TB1"
      Top             =   480
      Width           =   5655
   End
   Begin VB.CommandButton CMD1 
      BackColor       =   &H80000004&
      Caption         =   "Kembali Ke Menu Sebelum"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1245
      Left            =   2880
      MaskColor       =   &H00400000&
      Picture         =   "Frm26.frx":3494
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Kembali Ke Menu Sebelum"
      Top             =   3120
      Width           =   3225
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "No. Telefon :"
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "** Anda perlu menutup menu ini dahulu sebelum boleh ke menu seterusnya."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   2760
      Width           =   7095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "** Jika anda ingin rekod pembelian ini disimpan , sila daftarkan pembeli ini sebagai pelanggan kedai ke dalam sistem."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   2040
      Width           =   7095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "** Maklumat ini akan dipaparkan dalam invoice/voucher pembelian. Data ini tidak akan disimpan sebagai rekod pembelian pembeli ini."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   7095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Sila isikan maklumat pembeli dalam ruangan di bawah."
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   7095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nama :"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   480
      Width           =   1335
   End
End
Attribute VB_Name = "Frm26"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMD1_Click()
'on error resume next
Frm26.Hide
End Sub
Private Sub CMD2_Click()
'on error resume next
Note = "Padamkan semua maklumat pembeli ini ?" & vbCrLf & _
        vbNullString & vbCrLf & _
        "Teruskan ?"

Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")

If Answer = vbYes Then
    If MDI_frm1.L5_Text = 4 Then Frm84.L27_Text = vbNullString 'Nama Pembeli
    If MDI_frm1.L5_Text = 3 Then Frm83.L36_Text = vbNullString 'Nama Pembeli
    
    Frm26.TB1 = vbNullString
    Frm26.TB2 = vbNullString
    
    Frm26.TB1.SetFocus
End If
End Sub

Private Sub TB1_Change()
'on error resume next
If MDI_frm1.L5_Text = "4" Or MDI_frm1.L5_Text = "5" Then
    Frm84.L27_Text = UCase(Frm26.TB1) 'Nama Pembeli
End If
If MDI_frm1.L5_Text = "3" Then Frm83.L36_Text = UCase(Frm26.TB1) 'Nama Pembeli
If MDI_frm1.L5_Text = "7" Then Frm87.L5_Text = UCase(Frm26.TB1) 'Nama Pembeli
If MDI_frm1.L5_Text = "8" Then Frm93.L35_Text = UCase(Frm26.TB1) 'Nama Pembeli
If MDI_frm1.L5_Text = "10" Then Frm92.L51_Text = UCase(Frm26.TB1) 'Nama Pembeli
End Sub
