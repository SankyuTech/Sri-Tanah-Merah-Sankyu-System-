VERSION 5.00
Begin VB.Form Frm4 
   Caption         =   "Maklumat Sistem"
   ClientHeight    =   4575
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9780
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Frm4.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4575
   ScaleWidth      =   9780
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label41 
      BackStyle       =   0  'Transparent
      Caption         =   "Versi AB"
      Height          =   255
      Left            =   240
      TabIndex        =   41
      Top             =   4200
      Width           =   1935
   End
   Begin VB.Label Label40 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   255
      Left            =   2160
      TabIndex        =   40
      Top             =   4200
      Width           =   105
   End
   Begin VB.Label L6_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "L6_Text"
      Height          =   255
      Left            =   2280
      TabIndex        =   39
      Top             =   4200
      Width           =   4200
   End
   Begin VB.Label L2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "COPYRIGHT © Sankyu System 2013"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   4800
      TabIndex        =   38
      Top             =   4200
      Width           =   4815
   End
   Begin VB.Label L5_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "L5_Text"
      Height          =   255
      Left            =   2280
      TabIndex        =   37
      Top             =   3960
      Width           =   4200
   End
   Begin VB.Label L3_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "L3_Text"
      Height          =   255
      Left            =   2280
      TabIndex        =   36
      Top             =   3480
      Width           =   4200
   End
   Begin VB.Label L2_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "L2_Text"
      Height          =   255
      Left            =   2280
      TabIndex        =   35
      Top             =   3240
      Width           =   4200
   End
   Begin VB.Label L1_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "L1_Text"
      Height          =   255
      Left            =   2280
      TabIndex        =   34
      Top             =   3000
      Width           =   4200
   End
   Begin VB.Label L4_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "L4_Text"
      Height          =   255
      Left            =   2280
      TabIndex        =   33
      Top             =   3720
      Width           =   4200
   End
   Begin VB.Label Label33 
      BackStyle       =   0  'Transparent
      Caption         =   "http://sankyutech-visualbasic.weebly.com"
      Height          =   255
      Left            =   2040
      MouseIcon       =   "Frm4.frx":0ECA
      MousePointer    =   99  'Custom
      TabIndex        =   32
      Top             =   2040
      Width           =   5000
   End
   Begin VB.Label Label32 
      BackStyle       =   0  'Transparent
      Caption         =   "SankyuSystem"
      Height          =   255
      Left            =   2040
      TabIndex        =   31
      Top             =   1800
      Width           =   5000
   End
   Begin VB.Label Label31 
      BackStyle       =   0  'Transparent
      Caption         =   "Point Of Sales System"
      Height          =   255
      Left            =   2040
      MouseIcon       =   "Frm4.frx":11D4
      MousePointer    =   99  'Custom
      TabIndex        =   30
      Top             =   1560
      Width           =   5000
   End
   Begin VB.Label Label30 
      BackStyle       =   0  'Transparent
      Caption         =   "sankyusystem@gmail.com"
      Height          =   255
      Left            =   2040
      TabIndex        =   29
      Top             =   1320
      Width           =   5000
   End
   Begin VB.Label Label29 
      BackStyle       =   0  'Transparent
      Caption         =   "+6010 - 900 4788"
      Height          =   255
      Left            =   2040
      TabIndex        =   28
      Top             =   1080
      Width           =   5000
   End
   Begin VB.Label Label28 
      BackStyle       =   0  'Transparent
      Caption         =   "Insan"
      Height          =   255
      Left            =   2040
      TabIndex        =   27
      Top             =   840
      Width           =   5000
   End
   Begin VB.Label Label27 
      BackStyle       =   0  'Transparent
      Caption         =   "Sankyu System"
      Height          =   255
      Left            =   2040
      TabIndex        =   26
      Top             =   600
      Width           =   5000
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   255
      Left            =   2160
      TabIndex        =   25
      Top             =   3960
      Width           =   105
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   255
      Left            =   2160
      TabIndex        =   24
      Top             =   3720
      Width           =   105
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   255
      Left            =   2160
      TabIndex        =   23
      Top             =   3480
      Width           =   105
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   255
      Left            =   2160
      TabIndex        =   22
      Top             =   3240
      Width           =   105
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   255
      Left            =   2160
      TabIndex        =   21
      Top             =   3000
      Width           =   105
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Maklumat Sistem"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   20
      Top             =   2520
      Width           =   3735
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   255
      Left            =   1920
      TabIndex        =   19
      Top             =   2040
      Width           =   100
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   255
      Left            =   1920
      TabIndex        =   18
      Top             =   1800
      Width           =   100
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   255
      Left            =   1920
      TabIndex        =   17
      Top             =   1560
      Width           =   100
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   255
      Left            =   1920
      TabIndex        =   16
      Top             =   1320
      Width           =   100
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   255
      Left            =   1920
      TabIndex        =   15
      Top             =   1080
      Width           =   100
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   255
      Left            =   1920
      TabIndex        =   14
      Top             =   840
      Width           =   100
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   255
      Left            =   1920
      TabIndex        =   13
      Top             =   600
      Width           =   100
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Maklumat Developer"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   12
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Versi AE"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   3960
      Width           =   1935
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Versi database"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   3480
      Width           =   1935
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Versi sistem"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Nama kedai"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Website"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Instagram"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Facebook"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "No. Telefon"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Person in charge"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Versi database image"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   3720
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Developer "
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   3000
      Left            =   6600
      Picture         =   "Frm4.frx":14DE
      Top             =   240
      Width           =   3000
   End
End
Attribute VB_Name = "Frm4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
                (ByVal hwnd As Long, _
                ByVal lpOperation As String, _
                ByVal lpFile As String, _
                ByVal lpParameters As String, _
                ByVal lpDirectory As String, _
                ByVal nShowCmd As Long) As Long
Private Sub BrowseTo(ByRef pstrURL As String)
'On Error Resume Next
' Opens users default web browser and navigates to the selected URL
Call ShellExecute(Me.hwnd, "Open", pstrURL, "", "", True)
End Sub
Private Sub Label31_Click()
'On Error Resume Next
BrowseTo "https://www.facebook.com/ExcelVisualBasicApplicationVba"
End Sub
Private Sub Label33_Click()
'On Error Resume Next
BrowseTo "http://sankyutech-visualbasic.weebly.com"
End Sub
