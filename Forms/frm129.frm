VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm129 
   Caption         =   "Report Trade In"
   ClientHeight    =   12600
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   22320
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
   LinkTopic       =   "frm129"
   MDIChild        =   -1  'True
   ScaleHeight     =   12600
   ScaleWidth      =   22320
   WindowState     =   2  'Maximized
   Begin VB.CheckBox CB4 
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
      Left            =   9480
      TabIndex        =   33
      Top             =   1005
      Width           =   200
   End
   Begin VB.CheckBox CB3 
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
      Left            =   9480
      TabIndex        =   31
      Top             =   765
      Width           =   200
   End
   Begin VB.CheckBox CB2 
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
      Left            =   9480
      TabIndex        =   29
      Top             =   525
      Width           =   200
   End
   Begin VB.CommandButton CMD3 
      BackColor       =   &H000080FF&
      Caption         =   "Report"
      Height          =   405
      Left            =   3360
      MaskColor       =   &H00400000&
      MouseIcon       =   "frm129.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   5
      ToolTipText     =   "Report"
      Top             =   1320
      Width           =   5625
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
      Left            =   16920
      TabIndex        =   4
      Top             =   4125
      Visible         =   0   'False
      Width           =   200
   End
   Begin VB.ComboBox CBB1 
      BackColor       =   &H00FFFFFF&
      DataField       =   "Supplier"
      Height          =   360
      ItemData        =   "frm129.frx":030A
      Left            =   2160
      List            =   "frm129.frx":030C
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   435
      Width           =   7005
   End
   Begin VB.ComboBox CBB2 
      BackColor       =   &H00FFFFFF&
      DataField       =   "Supplier"
      Height          =   360
      ItemData        =   "frm129.frx":030E
      Left            =   2160
      List            =   "frm129.frx":0310
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   795
      Width           =   7005
   End
   Begin VB.CommandButton CMD22 
      BackColor       =   &H00FFFFFF&
      Height          =   650
      Left            =   14160
      MaskColor       =   &H00400000&
      MouseIcon       =   "frm129.frx":0312
      MousePointer    =   99  'Custom
      Picture         =   "frm129.frx":061C
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Paparan seterusnya"
      Top             =   10995
      Width           =   1100
   End
   Begin VB.CommandButton CMD21 
      BackColor       =   &H00FFFFFF&
      Height          =   650
      Left            =   12960
      MaskColor       =   &H00400000&
      MouseIcon       =   "frm129.frx":0F42
      MousePointer    =   99  'Custom
      Picture         =   "frm129.frx":124C
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Paparan sebelumnya"
      Top             =   10995
      Width           =   1100
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   360
      Left            =   16665
      TabIndex        =   6
      Top             =   5055
      Visible         =   0   'False
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
      Format          =   109641728
      CurrentDate     =   41561
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   360
      Left            =   16665
      TabIndex        =   7
      Top             =   5415
      Visible         =   0   'False
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
      Format          =   109641728
      CurrentDate     =   41561
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   8565
      Left            =   240
      TabIndex        =   8
      ToolTipText     =   "Sila ""Double Click"" untuk menu seterusnya."
      Top             =   2355
      Width           =   15045
      _ExtentX        =   26538
      _ExtentY        =   15108
      _Version        =   393216
      Rows            =   1
      Cols            =   0
      FixedCols       =   0
      BackColor       =   16777215
      ForeColor       =   0
      BackColorFixed  =   8454016
      BackColorSel    =   -2147483643
      ForeColorSel    =   12582912
      BackColorBkg    =   16777215
      GridColor       =   0
      WordWrap        =   -1  'True
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
   Begin VB.Label L15_Text 
      BackColor       =   &H00C0C0FF&
      Caption         =   "L15_Text"
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   19920
      TabIndex        =   36
      Top             =   2160
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Shape Shape2 
      Height          =   1815
      Left            =   120
      Top             =   75
      Width           =   15135
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Report ini adalah dikhaskan bagi barang trade in SAHAJA."
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
      Height          =   360
      Left            =   9480
      TabIndex        =   35
      Top             =   240
      Width           =   8370
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Report Stok"
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   9720
      TabIndex        =   34
      Top             =   960
      Width           =   2610
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Report Jualan"
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   9720
      TabIndex        =   32
      Top             =   720
      Width           =   2610
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Report Belian"
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   9720
      TabIndex        =   30
      Top             =   480
      Width           =   2610
   End
   Begin VB.Label Label62 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tarikh Mula :"
      Height          =   300
      Left            =   14955
      TabIndex        =   27
      Top             =   5100
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label63 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tarikh Akhir "
      Height          =   300
      Left            =   14955
      TabIndex        =   26
      Top             =   5460
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Sila klik jika ingin melihat senarai rekod di dalam tempoh tarikh di bawah."
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   17160
      TabIndex        =   25
      Top             =   4080
      Visible         =   0   'False
      Width           =   8370
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Kategori Produk * "
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   435
      TabIndex        =   24
      Top             =   435
      Width           =   1695
   End
   Begin VB.Label L5_Text 
      BackColor       =   &H00C0C0FF&
      Caption         =   "L5_Text"
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   19920
      TabIndex        =   23
      Top             =   315
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label L6_Text 
      BackColor       =   &H00C0C0FF&
      Caption         =   "L6_Text"
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   19920
      TabIndex        =   22
      Top             =   675
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label L7_Text 
      BackColor       =   &H00C0C0FF&
      Caption         =   "L7_Text"
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   19920
      TabIndex        =   21
      Top             =   1035
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label L8_Text 
      BackColor       =   &H00C0C0FF&
      Caption         =   "L8_Text"
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   19920
      TabIndex        =   20
      Top             =   1395
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Purity * "
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   435
      TabIndex        =   19
      Top             =   795
      Width           =   1695
   End
   Begin VB.Label L9_Text 
      BackColor       =   &H00C0C0FF&
      Caption         =   "L9_Text"
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   19920
      TabIndex        =   18
      Top             =   1755
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label L14_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "Senarai report                       "
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
      Left            =   360
      TabIndex        =   17
      Top             =   2040
      Width           =   15495
   End
   Begin VB.Label L70_Text 
      BackColor       =   &H8000000C&
      Caption         =   "L70_Text"
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   3720
      TabIndex        =   16
      Top             =   11115
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label L69_Text 
      BackColor       =   &H8000000C&
      Caption         =   "L69_Text"
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   3720
      TabIndex        =   15
      Top             =   11475
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
      Left            =   11880
      TabIndex        =   14
      Top             =   10995
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
      Left            =   12480
      TabIndex        =   13
      Top             =   10995
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Bil :"
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
      Left            =   -600
      TabIndex        =   12
      Top             =   10995
      Width           =   1455
   End
   Begin VB.Label L10_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "L10_Text"
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
      Left            =   960
      TabIndex        =   11
      Top             =   10995
      Width           =   1455
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Berat :"
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
      Left            =   -600
      TabIndex        =   10
      Top             =   11235
      Width           =   1455
   End
   Begin VB.Label L11_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "L11_Text"
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
      Left            =   960
      TabIndex        =   9
      Top             =   11235
      Width           =   1455
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
      Left            =   10560
      TabIndex        =   28
      Top             =   10995
      Width           =   2295
   End
   Begin VB.Menu frm129_pm_menu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu frm129_sm_excel 
         Caption         =   "Export Excel"
      End
   End
End
Attribute VB_Name = "frm129"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CB2_Click()
'on error resume next
If frm129.CB2 = 1 Then
    frm129.CB3 = 0
    frm129.CB4 = 0
End If
End Sub
Private Sub CB3_Click()
'on error resume next
If frm129.CB3 = 1 Then
    frm129.CB2 = 0
    frm129.CB4 = 0
End If
End Sub
Private Sub CB4_Click()
'on error resume next
If frm129.CB4 = 1 Then
    frm129.CB3 = 0
    frm129.CB2 = 0
End If
End Sub

Private Sub CMD21_Click()
'on error resume next
Dim frm129_LM_CURR_PAGE As Double
Dim frm129_LM_TOTAL_PAGE As Double

frm129_LM_CURR_PAGE = 0
frm129_LM_TOTAL_PAGE = 0

If frm129.L67_Text <> vbNullString And IsNumeric(frm129.L67_Text) Then
    If frm129.L68_Text <> vbNullString And IsNumeric(frm129.L68_Text) Then
        frm129_LM_CURR_PAGE = frm129.L67_Text
        frm129_LM_TOTAL_PAGE = frm129.L68_Text
        
        If frm129_LM_CURR_PAGE <> 1 And frm129_LM_CURR_PAGE <> 0 Then
        
            GM_NEXT_PREV = 1 '0 : Next , 1 : Previous
                                
            Call frm129_report_trade_in_header
            
            If frm129.L15_Text = "0" Then
                Call frm129_report_trade_in_belian
            ElseIf frm129.L15_Text = "1" Then
                Call frm129_report_trade_in_jualan
            ElseIf frm129.L15_Text = "2" Then
                Call frm129_report_trade_in_stok
            End If
            
        End If

    End If
End If
End Sub
Private Sub CMD22_Click()
'on error resume next
Dim frm129_LM_CURR_PAGE As Double
Dim frm129_LM_TOTAL_PAGE As Double

frm129_LM_CURR_PAGE = 0
frm129_LM_TOTAL_PAGE = 0

If frm129.L67_Text <> vbNullString And IsNumeric(frm129.L67_Text) Then
    If frm129.L68_Text <> vbNullString And IsNumeric(frm129.L68_Text) Then
        frm129_LM_CURR_PAGE = frm129.L67_Text
        frm129_LM_TOTAL_PAGE = frm129.L68_Text
        
        If frm129_LM_CURR_PAGE < frm129_LM_TOTAL_PAGE Then
        
            GM_NEXT_PREV = 0 '0 : Next , 1 : Previous
            
            Call frm129_report_trade_in_header
            
            If frm129.L15_Text = "0" Then
                Call frm129_report_trade_in_belian
            ElseIf frm129.L15_Text = "1" Then
                Call frm129_report_trade_in_jualan
            ElseIf frm129.L15_Text = "2" Then
                Call frm129_report_trade_in_stok
            End If
            
        End If
    End If
End If
End Sub

Private Sub CMD3_Click()
'On Error Resume Next
If frm129.CB2 = 0 And frm129.CB3 = 0 And frm129.CB4 = 0 Then

    MsgBox "Sila buat pilihan jenis report.", vbExclamation, "Info"
    
    Exit Sub
    
End If

If frm129.CB1 = 1 Then 'Pilihan tarikh
    frm129.L5_Text = 1
Else
    frm129.L5_Text = 0
End If

If frm129.CBB1 = vbNullString Then

    MsgBox "Sila buat pilihan kategori.", vbInformation, "Info"
    
    Exit Sub
    
End If

If frm129.CBB2 = vbNullString Then

    MsgBox "Sila buat pilihan purity.", vbInformation, "Info"
    
    Exit Sub
    
End If

If frm129.CB2 = 1 Then
    frm129.L15_Text = 0
ElseIf frm129.CB3 = 1 Then
    frm129.L15_Text = 1
ElseIf frm129.CB4 = 1 Then
    frm129.L15_Text = 2
End If

frm129.L6_Text = frm129.DTPicker1 'Tarikh mula
frm129.L7_Text = frm129.DTPicker2 'Tarikh akhir

frm129.L8_Text = frm129.CBB1 'Kategori
frm129.L9_Text = frm129.CBB2 'Purity

frm129.L69_Text = -1 'Titik Pencarian Data
frm129.L70_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
frm129.L67_Text = 0 'Paparan Page ke-xxx
frm129.L68_Text = 0

GM_NEXT_PREV = 0

Call frm129_report_trade_in_header

If frm129.CB2 = 1 Then
    Call frm129_report_trade_in_belian
ElseIf frm129.CB3 = 1 Then
    Call frm129_report_trade_in_jualan
ElseIf frm129.CB4 = 1 Then
    Call frm129_report_trade_in_stok
End If
End Sub

Private Sub frm129_sm_excel_Click()
'on error resume next
Dim TM As Date
Dim TA As Date

LM_FOUND = 0
frm129_LM_No_ID = vbNullString

If frm129.MSFlexGrid1 <> vbNullString Then

    If IsNumeric(frm129.MSFlexGrid1) Then
    
        frm129_LM_No_ID = frm129.MSFlexGrid1.TextMatrix(frm129.MSFlexGrid1, 2) 'No. ID
        
        If frm129_LM_No_ID <> vbNullString Then

            Note = "Adakah anda ingin export semua data ini ke excel?" & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Sistem mungkin mengambil masa untuk export semua data ini." & vbCrLf & _
                    "Sila tunggu sehingga sistem selesai export data ini." & vbCrLf & _
                    "Teruskan?"
                    
            Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
            
            If Answer = vbYes Then

                If frm129.L15_Text = "0" Then
                    Call frm129_excel_trade_in_belian
                ElseIf frm129.L15_Text = "1" Then
                    Call frm129_excel_trade_in_jualan
                ElseIf frm129.L15_Text = "2" Then
                    Call frm129_excel_trade_in_stok
                End If
            
            End If
            
        End If
        
    End If
    
End If
End Sub

Private Sub MSFlexGrid1_DblClick()
'On Error Resume Next
frm129_LM_No_ID = vbNullString

If frm129.MSFlexGrid1 <> vbNullString Then

    If IsNumeric(frm129.MSFlexGrid1) Then
    
        frm129_LM_No_ID = frm129.MSFlexGrid1.TextMatrix(frm129.MSFlexGrid1, 2) 'No. ID
        
        If frm129_LM_No_ID <> vbNullString Then
    
            PopupMenu frm129_pm_menu
            
        Else
            
            MsgBox "Tiada data.", vbExclamation, "Info"
            
        End If
        
    Else
    
        MsgBox "Tiada data.", vbExclamation, "Info"
        
    End If
    
End If
End Sub
