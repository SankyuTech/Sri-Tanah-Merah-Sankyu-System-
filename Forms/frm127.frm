VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frm127 
   Caption         =   "Log"
   ClientHeight    =   13035
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   23880
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
   Icon            =   "frm127.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   13035
   ScaleWidth      =   23880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Log"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   11655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   19575
      Begin VB.CommandButton CMD22 
         Caption         =   "Next"
         Height          =   810
         Left            =   18360
         MouseIcon       =   "frm127.frx":0ECA
         MousePointer    =   99  'Custom
         Picture         =   "frm127.frx":11D4
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Tutup senarai ini."
         Top             =   10800
         Width           =   1095
      End
      Begin VB.CommandButton CMD21 
         Caption         =   "Back"
         Height          =   810
         Left            =   17160
         MouseIcon       =   "frm127.frx":229E
         MousePointer    =   99  'Custom
         Picture         =   "frm127.frx":25A8
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Tutup senarai ini."
         Top             =   10800
         Width           =   1095
      End
      Begin VB.TextBox TB1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   1665
         TabIndex        =   14
         Text            =   "TB1"
         Top             =   1560
         Width           =   5655
      End
      Begin VB.ComboBox CBB3 
         BackColor       =   &H00FFFFFF&
         DataField       =   "Supplier"
         Height          =   360
         ItemData        =   "frm127.frx":3672
         Left            =   8745
         List            =   "frm127.frx":3674
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   960
         Width           =   5565
      End
      Begin VB.CheckBox CB1 
         BackColor       =   &H8000000C&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   200
         Left            =   280
         TabIndex        =   10
         Top             =   320
         Width           =   200
      End
      Begin VB.ComboBox CBB1 
         BackColor       =   &H00FFFFFF&
         DataField       =   "Supplier"
         Height          =   360
         ItemData        =   "frm127.frx":3676
         Left            =   8745
         List            =   "frm127.frx":3678
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   600
         Width           =   5565
      End
      Begin VB.CommandButton CMD1 
         Caption         =   "Senarai Log"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Left            =   14400
         MaskColor       =   &H00400000&
         MouseIcon       =   "frm127.frx":367A
         MousePointer    =   99  'Custom
         Picture         =   "frm127.frx":3984
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   3345
      End
      Begin VB.ComboBox CBB2 
         BackColor       =   &H00FFFFFF&
         DataField       =   "Supplier"
         Height          =   360
         ItemData        =   "frm127.frx":5F4E
         Left            =   8745
         List            =   "frm127.frx":5F50
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   5565
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   360
         Left            =   1665
         TabIndex        =   4
         Top             =   600
         Width           =   5550
         _ExtentX        =   9790
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
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   360
         Left            =   1665
         TabIndex        =   5
         Top             =   960
         Width           =   5550
         _ExtentX        =   9790
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
      Begin MSComctlLib.ListView LV1 
         Height          =   8370
         Left            =   120
         TabIndex        =   16
         Top             =   2280
         Width           =   19305
         _ExtentX        =   34052
         _ExtentY        =   14764
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label L12_Text 
         BackColor       =   &H00C0C0FF&
         Caption         =   "L12_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   9480
         TabIndex        =   31
         Top             =   1680
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label L11_Text 
         BackColor       =   &H00C0C0FF&
         Caption         =   "L11_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   9480
         TabIndex        =   30
         Top             =   1320
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label L10_Text 
         BackColor       =   &H00C0C0FF&
         Caption         =   "L10_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   8640
         TabIndex        =   29
         Top             =   2040
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label L9_Text 
         BackColor       =   &H00C0C0FF&
         Caption         =   "L9_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   8640
         TabIndex        =   28
         Top             =   1680
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label L8_Text 
         BackColor       =   &H00C0C0FF&
         Caption         =   "L8_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   8640
         TabIndex        =   27
         Top             =   1320
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label L7_Text 
         BackColor       =   &H00C0C0FF&
         Caption         =   "L7_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   7680
         TabIndex        =   26
         Top             =   2040
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label L6_Text 
         BackColor       =   &H00C0C0FF&
         Caption         =   "L6_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   7680
         TabIndex        =   25
         Top             =   1680
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label L5_Text 
         BackColor       =   &H00C0C0FF&
         Caption         =   "L5_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   7680
         TabIndex        =   24
         Top             =   1320
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label L70_Text 
         BackColor       =   &H8000000C&
         Caption         =   "L70_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   1920
         TabIndex        =   21
         Top             =   10920
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label L69_Text 
         BackColor       =   &H8000000C&
         Caption         =   "L69_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   1920
         TabIndex        =   20
         Top             =   11280
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label L67_Text 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "L67_Text"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   15960
         TabIndex        =   19
         Top             =   10800
         Width           =   375
      End
      Begin VB.Label L68_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L68_Text"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   16560
         TabIndex        =   18
         Top             =   10800
         Width           =   615
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Paparan Muka  :          / "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   14640
         TabIndex        =   17
         Top             =   10800
         Width           =   2295
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Keyword :"
         Height          =   300
         Left            =   120
         TabIndex        =   15
         Top             =   1560
         Width           =   1500
      End
      Begin VB.Shape Shape1 
         Height          =   1215
         Left            =   120
         Top             =   240
         Width           =   7215
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "User * :"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   7200
         TabIndex        =   13
         Top             =   990
         Width           =   1500
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Sila klik jika ingin melihat senarai log di dalam tempoh tarikh di bawah."
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   520
         TabIndex        =   11
         Top             =   280
         Width           =   8370
      End
      Begin VB.Label Label62 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tarikh Mula * :"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   645
         Width           =   1500
      End
      Begin VB.Label Label63 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tarikh Akhir * :"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1005
         Width           =   1500
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Terminal * :"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   7200
         TabIndex        =   7
         Top             =   630
         Width           =   1500
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cawangan * :"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   7200
         TabIndex        =   6
         Top             =   270
         Width           =   1500
      End
   End
End
Attribute VB_Name = "frm127"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMD1_Click()
'on error resume next
If InStr(1, frm127.TB1, "*") <> 0 Or InStr(1, frm127.TB1, "/") <> 0 Or InStr(1, frm127.TB1, "\") <> 0 Or InStr(1, frm127.TB1, "'") <> 0 Then

    MsgBox "Keyword lama mengandungi simbol yang tidak dibenarkan.", vbExclamation, "Error"
    Exit Sub
    
End If

If frm127.CB1 = 1 Then '0 : Tiada pilihan tarikh , 1 : Ada pilihan tarikh
    frm127.L5_Text = 1 '0 : Tiada pilihan tarikh , 1 : Ada pilihan tarikh
Else
    frm127.L5_Text = 0
End If

frm127.L6_Text = frm127.DTPicker1 'Tarikh Mula
frm127.L7_Text = frm127.DTPicker2 'Tarikh Akhir
frm127.L8_Text = frm127.TB1 'Keyword
frm127.L9_Text = frm127.CBB2 'Cawangan
frm127.L10_Text = frm127.CBB1 'Terminal
frm127.L11_Text = frm127.CBB3 'User

frm127.L69_Text = -1 'Titik Pencarian Data
frm127.L70_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
frm127.L67_Text = 0 'Paparan Page ke-xxx
frm127.L68_Text = 0

GM_NEXT_PREV = 0

Call frm127_log_header
Call frm127_log
End Sub

Private Sub CMD21_Click()
'on error resume next
Dim frm127_LM_CURR_PAGE As Double
Dim frm127_LM_TOTAL_PAGE As Double

frm127_LM_CURR_PAGE = 0
frm127_LM_TOTAL_PAGE = 0

If frm127.L67_Text <> vbNullString And IsNumeric(frm127.L67_Text) Then
    If frm127.L68_Text <> vbNullString And IsNumeric(frm127.L68_Text) Then
        frm127_LM_CURR_PAGE = frm127.L67_Text
        frm127_LM_TOTAL_PAGE = frm127.L68_Text
        
        If frm127_LM_CURR_PAGE <> 1 And frm127_LM_CURR_PAGE <> 0 Then
        
            GM_NEXT_PREV = 1 '0 : Next , 1 : Previous
                    
            Call frm127_log_header
            Call frm127_log
            
        End If

    End If
End If
End Sub
Private Sub CMD22_Click()
'on error resume next
Dim frm127_LM_CURR_PAGE As Double
Dim frm127_LM_TOTAL_PAGE As Double

frm127_LM_CURR_PAGE = 0
frm127_LM_TOTAL_PAGE = 0

If frm127.L67_Text <> vbNullString And IsNumeric(frm127.L67_Text) Then
    If frm127.L68_Text <> vbNullString And IsNumeric(frm127.L68_Text) Then
        frm127_LM_CURR_PAGE = frm127.L67_Text
        frm127_LM_TOTAL_PAGE = frm127.L68_Text
        
        If frm127_LM_CURR_PAGE < frm127_LM_TOTAL_PAGE Then
        
            GM_NEXT_PREV = 0 '0 : Next , 1 : Previous
            
            Call frm127_log_header
            Call frm127_log
                        
        End If
    End If
End If
End Sub

