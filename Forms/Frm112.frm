VERSION 5.00
Begin VB.Form Frm112 
   Caption         =   "Setting default printer"
   ClientHeight    =   9960
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14490
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
   Icon            =   "Frm112.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9960
   ScaleWidth      =   14490
   WindowState     =   2  'Maximized
   Begin VB.ListBox List1 
      Height          =   2700
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   7575
   End
   Begin VB.CommandButton CMD1 
      Caption         =   "Simpan tetapan printer"
      Height          =   375
      Left            =   2280
      MouseIcon       =   "Frm112.frx":0ECA
      MousePointer    =   99  'Custom
      TabIndex        =   0
      ToolTipText     =   "Penerimaan stok baru"
      Top             =   3360
      Width           =   2175
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Senarai printer yang dijumpai di dalam komputer ini."
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
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   6705
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Note *"
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
      Left            =   240
      TabIndex        =   3
      Top             =   3720
      Width           =   6705
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"Frm112.frx":11D4
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   3405
      Left            =   240
      TabIndex        =   2
      Top             =   3960
      Width           =   6705
   End
End
Attribute VB_Name = "Frm112"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit

'To verify that the printer has been changed to the default, open your "Printers" folder
'in your control panel and select a printer and press the right mouse button to access
'the context menu. Take notice of "Set As Default" check mark for the selected printer
'when your run this app...

Private cSetPrinter As New cSetDfltPrinter
Private Sub CMD1_Click()
'On Error Resume Next
Dim sMsg As String
Dim DeviceName As String
Dim L_PRINTER As String

Printer_Selected = 0

If List1.SelCount = 1 Then
    DeviceName = List1.List(List1.ListIndex)
    If cSetPrinter.SetPrinterAsDefault(DeviceName) Then
    
        L_PRINTER = DeviceName
        Printer_Selected = 1
        
    Else
        sMsg = DeviceName & " tidak berjaya untuk dijadikan default printer. Sila keluar dari menu ini dan cuba sekali lagi."
        MsgBox sMsg, vbExclamation, App.Title
    End If
    
Else
    MsgBox "Sila pilih printer yang ingin dijadikan default printer.", vbInformation, App.Title
End If

If Printer_Selected = 1 Then

    Note = "Adakah anda ingin tetapkan " & L_PRINTER & " sebagai default printer bagi sistem?"
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbNo Then
        Exit Sub
    End If
    If Answer = vbYes Then

        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from default_setting where Default1='" & "default" & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
        
            rs!default_printer = L_PRINTER
            rs.Update
    
        End If
        
        rs.Close
        Set rs = Nothing
    
        sMsg = L_PRINTER & " telah berjaya dijadikan sebagai default printer."
        
        MsgBox sMsg, vbExclamation, App.Title
        
    End If
        
End If
End Sub
Private Sub Form_Load()
'On Error Resume Next
Frm112.Picture = MDI_frm1.Picture

Dim i As Integer

For i = 0 To Printers.Count - 1
    List1.AddItem Printers(i).DeviceName
Next i
End Sub

