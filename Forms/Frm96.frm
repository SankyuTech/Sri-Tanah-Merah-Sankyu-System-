VERSION 5.00
Begin VB.Form Frm96 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Pilihan Cawangan"
   ClientHeight    =   1980
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11340
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
   Icon            =   "Frm96.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1980
   ScaleWidth      =   11340
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CMD2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Teruskan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6720
      MouseIcon       =   "Frm96.frx":0ECA
      MousePointer    =   99  'Custom
      Picture         =   "Frm96.frx":11D4
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1000
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.ComboBox CBB1 
      BackColor       =   &H00FFFFFF&
      DataField       =   "Supplier"
      Height          =   360
      ItemData        =   "Frm96.frx":229E
      Left            =   5040
      List            =   "Frm96.frx":22A0
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   600
      Width           =   6015
   End
   Begin VB.CommandButton CMD1 
      Caption         =   "Teruskan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6720
      MouseIcon       =   "Frm96.frx":22A2
      MousePointer    =   99  'Custom
      Picture         =   "Frm96.frx":25AC
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1000
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   2400
      Left            =   120
      Picture         =   "Frm96.frx":3676
      Top             =   120
      Width           =   3330
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cawangan * :"
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
      Height          =   330
      Left            =   3600
      TabIndex        =   2
      Top             =   615
      Width           =   1335
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Sila pilih cawangan."
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
      Height          =   285
      Left            =   3600
      TabIndex        =   0
      Top             =   240
      Width           =   6705
   End
End
Attribute VB_Name = "Frm96"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMD1_Click()
'on error resume next
LM_FOUND = 0

If Frm96.CBB1 = vbNullString Then

    MsgBox "Sila pilih cawangan", vbExclamation, "Info"
    
    Exit Sub
    
End If

MDI_frm1.L20_Text = Frm96.CBB1
If MDI_frm1.L20_Text <> "Semua cawangan" Then G_KEDAI = MDI_frm1.L20_Text

Call main_setting_kedai
Call main_setting

'If LM_FOUND = 1 Then

    Unload Frm96
    
    MsgBox "Anda sekarang sedang akses ke dalam data cawangan [" & UCase(G_CAWANGAN) & "].", vbInformation, "Sankyu System"
    
'Else
    
'    MsgBox "Ralat telah berlaku. Sila restart sistem!", vbCritical, "Info"

'End If
End Sub

Private Sub CMD2_Click()
'on error resume next
If Frm96.CBB1 = vbNullString Then

    MsgBox "Sila pilih cawangan", vbExclamation, "Info"
    Exit Sub
    
End If

If Frm96.CBB1 = "Semua cawangan" Then

    MsgBox "Pilihan Semua cawangan tidak dibenarkan. Sila pilih cawangan.", vbExclamation, "Info"
    Exit Sub
    
End If

MDI_frm1.L20_Text = Frm96.CBB1
G_KEDAI = Frm96.CBB1

Call main_setting_kedai
Call main_setting

'If LM_FOUND = 1 Then

    Unload Frm96
    MDI_frm1.CMD44.Enabled = False
    
    MsgBox "Anda sekarang sedang akses ke dalam data cawangan [" & UCase(G_CAWANGAN) & "].", vbInformation, "Sankyu System"
    
'Else
    
'    MsgBox "Ralat telah berlaku. Sila restart sistem!", vbCritical, "Info"

'End If
End Sub
