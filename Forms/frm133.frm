VERSION 5.00
Begin VB.Form frm133 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Pilihan Dulang"
   ClientHeight    =   2310
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9930
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
   ScaleHeight     =   2310
   ScaleWidth      =   9930
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CMD2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Batal"
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
      Left            =   5760
      MouseIcon       =   "frm133.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "frm133.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton CMD1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pilih"
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
      Left            =   4200
      MouseIcon       =   "frm133.frx":13D4
      MousePointer    =   99  'Custom
      Picture         =   "frm133.frx":16DE
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1320
      Width           =   1455
   End
   Begin VB.ComboBox CBB1 
      BackColor       =   &H00FFFFFF&
      DataField       =   "Supplier"
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
      ItemData        =   "frm133.frx":27A8
      Left            =   2760
      List            =   "frm133.frx":27AA
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   480
      Width           =   7000
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Dulang * :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   540
      Left            =   480
      TabIndex        =   1
      Top             =   480
      Width           =   2295
   End
End
Attribute VB_Name = "frm133"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMD1_Click()
'on error resume next
frm132.L1_Text = frm133.CBB1
Unload frm133
frm132.TB1.SetFocus
End Sub
Private Sub CMD2_Click()
'on error resume next
Unload frm133
End Sub
