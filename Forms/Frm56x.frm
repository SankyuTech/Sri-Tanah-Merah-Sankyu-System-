VERSION 5.00
Begin VB.Form Frm56x 
   Caption         =   "Tetapan Barcode"
   ClientHeight    =   12930
   ClientLeft      =   120
   ClientTop       =   450
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
   Icon            =   "Frm56x.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   12930
   ScaleWidth      =   23760
   WindowState     =   2  'Maximized
   Begin VB.CheckBox CB14 
      Caption         =   "Scanner Mode"
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
      Left            =   360
      TabIndex        =   60
      Top             =   3525
      Width           =   200
   End
   Begin VB.CheckBox CB15 
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
      Left            =   2835
      TabIndex        =   59
      Top             =   3525
      Width           =   200
   End
   Begin VB.ComboBox CBB16 
      Height          =   360
      Left            =   8880
      Style           =   2  'Dropdown List
      TabIndex        =   54
      Top             =   2280
      Width           =   3135
   End
   Begin VB.ComboBox CBB15 
      Height          =   360
      Left            =   8880
      Style           =   2  'Dropdown List
      TabIndex        =   52
      Top             =   1800
      Width           =   3135
   End
   Begin VB.ComboBox CBB14 
      Height          =   360
      Left            =   8880
      Style           =   2  'Dropdown List
      TabIndex        =   50
      Top             =   870
      Width           =   3135
   End
   Begin VB.ComboBox CBB13 
      Height          =   360
      Left            =   15000
      Style           =   2  'Dropdown List
      TabIndex        =   49
      Top             =   2280
      Width           =   1095
   End
   Begin VB.ComboBox CBB12 
      Height          =   360
      Left            =   15000
      Style           =   2  'Dropdown List
      TabIndex        =   48
      Top             =   1800
      Width           =   1095
   End
   Begin VB.ComboBox CBB11 
      Height          =   360
      Left            =   15000
      Style           =   2  'Dropdown List
      TabIndex        =   47
      Top             =   1350
      Width           =   1095
   End
   Begin VB.ComboBox CBB10 
      Height          =   360
      Left            =   15000
      Style           =   2  'Dropdown List
      TabIndex        =   42
      Top             =   870
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   3495
      Left            =   360
      ScaleHeight     =   3495
      ScaleWidth      =   6015
      TabIndex        =   35
      Top             =   8880
      Visible         =   0   'False
      Width           =   6015
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "RT001013"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   41
         Top             =   1320
         Width           =   5415
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Label11"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   2400
         Width           =   5415
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Label11"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   2160
         Width           =   5415
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Label11"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   1920
         Width           =   5415
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Label11"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   1680
         Width           =   5415
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Preview :"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   120
         TabIndex        =   36
         Top             =   120
         Width           =   1680
      End
      Begin VB.Image Image1 
         Height          =   750
         Left            =   120
         Picture         =   "Frm56x.frx":0ECA
         Top             =   480
         Width           =   2790
      End
   End
   Begin VB.CommandButton CMD1 
      Caption         =   "Simpan Tetapan"
      Height          =   375
      Left            =   5880
      MouseIcon       =   "Frm56x.frx":1B54
      MousePointer    =   99  'Custom
      TabIndex        =   34
      ToolTipText     =   "Senarai barang yang telah dimasukkan ke dalam senarai belian"
      Top             =   3480
      Width           =   3375
   End
   Begin VB.CheckBox CB4 
      Caption         =   "Scanner Mode"
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
      Left            =   120
      TabIndex        =   32
      Top             =   2355
      Width           =   200
   End
   Begin VB.ComboBox CBB9 
      Height          =   360
      Left            =   5400
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   2280
      Width           =   3135
   End
   Begin VB.ComboBox CBB8 
      Height          =   360
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   2280
      Width           =   3135
   End
   Begin VB.CheckBox CB3 
      Caption         =   "Scanner Mode"
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
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      Width           =   200
   End
   Begin VB.ComboBox CBB7 
      Height          =   360
      Left            =   5400
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1800
      Width           =   3135
   End
   Begin VB.ComboBox CBB6 
      Height          =   360
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1830
      Width           =   3135
   End
   Begin VB.ComboBox CBB5 
      Height          =   360
      Left            =   8880
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1350
      Width           =   3135
   End
   Begin VB.CheckBox CB2 
      Caption         =   "Scanner Mode"
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
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   200
   End
   Begin VB.ComboBox CBB4 
      Height          =   360
      Left            =   5400
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1350
      Width           =   3135
   End
   Begin VB.ComboBox CBB3 
      Height          =   360
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1350
      Width           =   3135
   End
   Begin VB.ComboBox CBB2 
      Height          =   360
      Left            =   5400
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   870
      Width           =   3135
   End
   Begin VB.CheckBox CB1 
      Caption         =   "Scanner Mode"
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
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   200
   End
   Begin VB.ComboBox CBB1 
      Height          =   360
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   870
      Width           =   3135
   End
   Begin VB.Timer Tmr1 
      Interval        =   10
      Left            =   1200
      Top             =   0
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "User perlu log out dari sistem (Semua station) bagi memboleh tetapan ini diupdate selepas tetapan ini disimpan."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   240
      TabIndex        =   64
      Top             =   4080
      Width           =   12135
   End
   Begin VB.Shape Shape1 
      Height          =   975
      Left            =   120
      Top             =   2880
      Width           =   5535
   End
   Begin VB.Label Label44 
      BackStyle       =   0  'Transparent
      Caption         =   "Jenis Barcode Label"
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
      Height          =   315
      Left            =   240
      TabIndex        =   63
      Top             =   2940
      Width           =   2295
   End
   Begin VB.Label Label46 
      BackStyle       =   0  'Transparent
      Caption         =   "Type A (35mm X 25mm)      Type B (75mm X 35mm)"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   600
      TabIndex        =   62
      Top             =   3480
      Width           =   4815
   End
   Begin VB.Label Label52 
      BackStyle       =   0  'Transparent
      Caption         =   "Jenis barcode label yang digunakan."
      Height          =   255
      Left            =   240
      TabIndex        =   61
      Top             =   3240
      Width           =   7935
   End
   Begin VB.Label L12_Text 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   3000
      TabIndex        =   58
      Top             =   8400
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Label L13_Text 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   6720
      TabIndex        =   57
      Top             =   8400
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Label L14_Text 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   10320
      TabIndex        =   56
      Top             =   8400
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "/"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   8520
      TabIndex        =   55
      Top             =   2310
      Width           =   360
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "/"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   8520
      TabIndex        =   53
      Top             =   1875
      Width           =   360
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "/"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   8520
      TabIndex        =   51
      Top             =   915
      Width           =   360
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Saiz tulisan barisan keempat"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   12120
      TabIndex        =   46
      Top             =   2310
      Width           =   3000
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Saiz tulisan barisan ketiga"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   12120
      TabIndex        =   45
      Top             =   1875
      Width           =   3000
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Saiz tulisan barisan kedua"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   12120
      TabIndex        =   44
      Top             =   1395
      Width           =   3000
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Saiz tulisan barisan pertama"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   12120
      TabIndex        =   43
      Top             =   915
      Width           =   3000
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Barisan Keempat"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   360
      TabIndex        =   33
      Top             =   2310
      Width           =   1680
   End
   Begin VB.Label L11_Text 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   10320
      TabIndex        =   31
      Top             =   7920
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Label L10_Text 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   6720
      TabIndex        =   30
      Top             =   7920
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Label L9_Text 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   3000
      TabIndex        =   29
      Top             =   7920
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Label L8_Text 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   10320
      TabIndex        =   28
      Top             =   7440
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Label L7_Text 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   6720
      TabIndex        =   27
      Top             =   7440
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Label L6_Text 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   3000
      TabIndex        =   26
      Top             =   7440
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Label L5_Text 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   10320
      TabIndex        =   25
      Top             =   6960
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Label L4_Text 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   6720
      TabIndex        =   24
      Top             =   6960
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Label L3_Text 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   3000
      TabIndex        =   23
      Top             =   6960
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "/"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   5040
      TabIndex        =   22
      Top             =   2310
      Width           =   360
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "/"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   5040
      TabIndex        =   21
      Top             =   1875
      Width           =   360
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Barisan Ketiga "
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   360
      TabIndex        =   20
      Top             =   1875
      Width           =   1680
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "/"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   5040
      TabIndex        =   19
      Top             =   1395
      Width           =   360
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Barisan Kedua"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   360
      TabIndex        =   18
      Top             =   1395
      Width           =   1680
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "/"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   8520
      TabIndex        =   17
      Top             =   1395
      Width           =   360
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "/"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   5040
      TabIndex        =   16
      Top             =   915
      Width           =   360
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Barisan Pertama"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   360
      TabIndex        =   15
      Top             =   915
      Width           =   1560
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sila buat penetapan barcode yang akan dicetak pada setiap barang kemas."
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   120
      TabIndex        =   14
      Top             =   240
      Width           =   9240
   End
   Begin VB.Label L2_Text 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "88/88/8888"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   21135
      TabIndex        =   13
      Top             =   435
      Visible         =   0   'False
      Width           =   2100
   End
   Begin VB.Label L1_Text 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "88/88/8888"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   21135
      TabIndex        =   12
      Top             =   120
      Visible         =   0   'False
      Width           =   2100
   End
End
Attribute VB_Name = "Frm56x"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CB14_Click()
'On Error Resume Next
If Frm56x.CB14 = 1 Then
    Frm56x.CB15 = 0
End If
End Sub
Private Sub CB15_Click()
'On Error Resume Next
If Frm56x.CB15 = 1 Then
    Frm56x.CB14 = 0
End If
End Sub
Private Sub CBB1_Click()
'On Error Resume Next
If Frm56x.CBB1 = "Berat" Then
    Frm56x.L3_Text = "BARCODE_BERAT"
ElseIf Frm56x.CBB1 = "Upah Modal" Then
    Frm56x.L3_Text = "BARCODE_UPAH"
ElseIf Frm56x.CBB1 = "Upah Jualan" Then
    Frm56x.L3_Text = "BARCODE_UPAH2"
ElseIf Frm56x.CBB1 = "Purity" Then
    Frm56x.L3_Text = "BARCODE_PURITY"
ElseIf Frm56x.CBB1 = "Panjang" Then
    Frm56x.L3_Text = "BARCODE_Panjang"
ElseIf Frm56x.CBB1 = "Lebar" Then
    Frm56x.L3_Text = "BARCODE_Lebar"
ElseIf Frm56x.CBB1 = "Saiz" Then
    Frm56x.L3_Text = "BARCODE_Saiz"
ElseIf Frm56x.CBB1 = "Dulang" Then
    Frm56x.L3_Text = "BARCODE_DULANG"
ElseIf Frm56x.CBB1 = "Supplier" Then
    Frm56x.L3_Text = "BARCODE_SUPPLIER"
ElseIf Frm56x.CBB1 = "Code 1" Then
    Frm56x.L3_Text = "BARCODE_CODE1"
ElseIf Frm56x.CBB1 = "Code 2" Then
    Frm56x.L3_Text = "BARCODE_CODE2"
ElseIf Frm56x.CBB1 = "Berat Riyal" Then
'ElseIf Frm56x.CBB1 = "Berat Amah" Then
    Frm56x.L3_Text = "BARCODE_RIYAL"
Else
    Frm56x.L3_Text = vbNullString
End If
End Sub
Private Sub CBB2_Click()
'On Error Resume Next
If Frm56x.CBB2 = "Berat" Then
    Frm56x.L4_Text = "BARCODE_BERAT"
ElseIf Frm56x.CBB2 = "Upah Modal" Then
    Frm56x.L4_Text = "BARCODE_UPAH"
ElseIf Frm56x.CBB2 = "Upah Jualan" Then
    Frm56x.L4_Text = "BARCODE_UPAH2"
ElseIf Frm56x.CBB2 = "Purity" Then
    Frm56x.L4_Text = "BARCODE_PURITY"
ElseIf Frm56x.CBB2 = "Panjang" Then
    Frm56x.L4_Text = "BARCODE_Panjang"
ElseIf Frm56x.CBB2 = "Lebar" Then
    Frm56x.L4_Text = "BARCODE_Lebar"
ElseIf Frm56x.CBB2 = "Saiz" Then
    Frm56x.L4_Text = "BARCODE_Saiz"
ElseIf Frm56x.CBB2 = "Dulang" Then
    Frm56x.L4_Text = "BARCODE_DULANG"
ElseIf Frm56x.CBB2 = "Supplier" Then
    Frm56x.L4_Text = "BARCODE_SUPPLIER"
ElseIf Frm56x.CBB2 = "Code 1" Then
    Frm56x.L4_Text = "BARCODE_CODE1"
ElseIf Frm56x.CBB2 = "Code 2" Then
    Frm56x.L4_Text = "BARCODE_CODE2"
ElseIf Frm56x.CBB2 = "Berat Riyal" Then
'ElseIf Frm56x.CBB2 = "Berat Amah" Then
    Frm56x.L4_Text = "BARCODE_RIYAL"
Else
    Frm56x.L4_Text = vbNullString
End If
End Sub
Private Sub CBB3_Click()
'On Error Resume Next
If Frm56x.CBB3 = "Berat" Then
    Frm56x.L5_Text = "BARCODE_BERAT"
ElseIf Frm56x.CBB3 = "Upah Modal" Then
    Frm56x.L5_Text = "BARCODE_UPAH"
ElseIf Frm56x.CBB3 = "Upah Jualan" Then
    Frm56x.L5_Text = "BARCODE_UPAH2"
ElseIf Frm56x.CBB3 = "Purity" Then
    Frm56x.L5_Text = "BARCODE_PURITY"
ElseIf Frm56x.CBB3 = "Panjang" Then
    Frm56x.L5_Text = "BARCODE_Panjang"
ElseIf Frm56x.CBB3 = "Lebar" Then
    Frm56x.L5_Text = "BARCODE_Lebar"
ElseIf Frm56x.CBB3 = "Saiz" Then
    Frm56x.L5_Text = "BARCODE_Saiz"
ElseIf Frm56x.CBB3 = "Dulang" Then
    Frm56x.L5_Text = "BARCODE_DULANG"
ElseIf Frm56x.CBB3 = "Supplier" Then
    Frm56x.L5_Text = "BARCODE_SUPPLIER"
ElseIf Frm56x.CBB3 = "Code 1" Then
    Frm56x.L5_Text = "BARCODE_CODE1"
ElseIf Frm56x.CBB3 = "Code 2" Then
    Frm56x.L5_Text = "BARCODE_CODE2"
ElseIf Frm56x.CBB3 = "Berat Riyal" Then
'ElseIf Frm56x.CBB3 = "Berat Amah" Then
    Frm56x.L5_Text = "BARCODE_RIYAL"
Else
    Frm56x.L5_Text = vbNullString
End If
End Sub
Private Sub CBB4_Click()
'On Error Resume Next
If Frm56x.CBB4 = "Berat" Then
    Frm56x.L6_Text = "BARCODE_BERAT"
ElseIf Frm56x.CBB4 = "Upah Modal" Then
    Frm56x.L6_Text = "BARCODE_UPAH"
ElseIf Frm56x.CBB4 = "Upah Jualan" Then
    Frm56x.L6_Text = "BARCODE_UPAH2"
ElseIf Frm56x.CBB4 = "Purity" Then
    Frm56x.L6_Text = "BARCODE_PURITY"
ElseIf Frm56x.CBB4 = "Panjang" Then
    Frm56x.L6_Text = "BARCODE_Panjang"
ElseIf Frm56x.CBB4 = "Lebar" Then
    Frm56x.L6_Text = "BARCODE_Lebar"
ElseIf Frm56x.CBB4 = "Saiz" Then
    Frm56x.L6_Text = "BARCODE_Saiz"
ElseIf Frm56x.CBB4 = "Dulang" Then
    Frm56x.L6_Text = "BARCODE_DULANG"
ElseIf Frm56x.CBB4 = "Supplier" Then
    Frm56x.L6_Text = "BARCODE_SUPPLIER"
ElseIf Frm56x.CBB4 = "Code 1" Then
    Frm56x.L6_Text = "BARCODE_CODE1"
ElseIf Frm56x.CBB4 = "Code 2" Then
    Frm56x.L6_Text = "BARCODE_CODE2"
ElseIf Frm56x.CBB4 = "Berat Riyal" Then
'ElseIf Frm56x.CBB4 = "Berat Amah" Then
    Frm56x.L6_Text = "BARCODE_RIYAL"
Else
    Frm56x.L6_Text = vbNullString
End If
End Sub
Private Sub CBB5_Click()
'On Error Resume Next
If Frm56x.CBB5 = "Berat" Then
    Frm56x.L7_Text = "BARCODE_BERAT"
ElseIf Frm56x.CBB5 = "Upah Modal" Then
    Frm56x.L7_Text = "BARCODE_UPAH"
ElseIf Frm56x.CBB5 = "Upah Jualan" Then
    Frm56x.L7_Text = "BARCODE_UPAH2"
ElseIf Frm56x.CBB5 = "Purity" Then
    Frm56x.L7_Text = "BARCODE_PURITY"
ElseIf Frm56x.CBB5 = "Panjang" Then
    Frm56x.L7_Text = "BARCODE_Panjang"
ElseIf Frm56x.CBB5 = "Lebar" Then
    Frm56x.L7_Text = "BARCODE_Lebar"
ElseIf Frm56x.CBB5 = "Saiz" Then
    Frm56x.L7_Text = "BARCODE_Saiz"
ElseIf Frm56x.CBB5 = "Dulang" Then
    Frm56x.L7_Text = "BARCODE_DULANG"
ElseIf Frm56x.CBB5 = "Supplier" Then
    Frm56x.L7_Text = "BARCODE_SUPPLIER"
ElseIf Frm56x.CBB5 = "Code 1" Then
    Frm56x.L7_Text = "BARCODE_CODE1"
ElseIf Frm56x.CBB5 = "Code 2" Then
    Frm56x.L7_Text = "BARCODE_CODE2"
ElseIf Frm56x.CBB5 = "Berat Riyal" Then
'ElseIf Frm56x.CBB5 = "Berat Amah" Then
    Frm56x.L7_Text = "BARCODE_RIYAL"
Else
    Frm56x.L7_Text = vbNullString
End If
End Sub
Private Sub CBB6_Click()
'On Error Resume Next
If Frm56x.CBB6 = "Berat" Then
    Frm56x.L8_Text = "BARCODE_BERAT"
ElseIf Frm56x.CBB6 = "Upah Modal" Then
    Frm56x.L8_Text = "BARCODE_UPAH"
ElseIf Frm56x.CBB6 = "Upah Jualan" Then
    Frm56x.L8_Text = "BARCODE_UPAH2"
ElseIf Frm56x.CBB6 = "Purity" Then
    Frm56x.L8_Text = "BARCODE_PURITY"
ElseIf Frm56x.CBB6 = "Panjang" Then
    Frm56x.L8_Text = "BARCODE_Panjang"
ElseIf Frm56x.CBB6 = "Lebar" Then
    Frm56x.L8_Text = "BARCODE_Lebar"
ElseIf Frm56x.CBB6 = "Saiz" Then
    Frm56x.L8_Text = "BARCODE_Saiz"
ElseIf Frm56x.CBB6 = "Dulang" Then
    Frm56x.L8_Text = "BARCODE_DULANG"
ElseIf Frm56x.CBB6 = "Supplier" Then
    Frm56x.L8_Text = "BARCODE_SUPPLIER"
ElseIf Frm56x.CBB6 = "Code 1" Then
    Frm56x.L8_Text = "BARCODE_CODE1"
ElseIf Frm56x.CBB6 = "Code 2" Then
    Frm56x.L8_Text = "BARCODE_CODE2"
ElseIf Frm56x.CBB6 = "Berat Riyal" Then
'ElseIf Frm56x.CBB6 = "Berat Amah" Then
    Frm56x.L8_Text = "BARCODE_RIYAL"
Else
    Frm56x.L8_Text = vbNullString
End If
End Sub
Private Sub CBB7_Click()
'On Error Resume Next
If Frm56x.CBB7 = "Berat" Then
    Frm56x.L9_Text = "BARCODE_BERAT"
ElseIf Frm56x.CBB7 = "Upah Modal" Then
    Frm56x.L9_Text = "BARCODE_UPAH"
ElseIf Frm56x.CBB7 = "Upah Jualan" Then
    Frm56x.L9_Text = "BARCODE_UPAH2"
ElseIf Frm56x.CBB7 = "Purity" Then
    Frm56x.L9_Text = "BARCODE_PURITY"
ElseIf Frm56x.CBB7 = "Panjang" Then
    Frm56x.L9_Text = "BARCODE_Panjang"
ElseIf Frm56x.CBB7 = "Lebar" Then
    Frm56x.L9_Text = "BARCODE_Lebar"
ElseIf Frm56x.CBB7 = "Saiz" Then
    Frm56x.L9_Text = "BARCODE_Saiz"
ElseIf Frm56x.CBB7 = "Dulang" Then
    Frm56x.L9_Text = "BARCODE_DULANG"
ElseIf Frm56x.CBB7 = "Supplier" Then
    Frm56x.L9_Text = "BARCODE_SUPPLIER"
ElseIf Frm56x.CBB7 = "Code 1" Then
    Frm56x.L9_Text = "BARCODE_CODE1"
ElseIf Frm56x.CBB7 = "Code 2" Then
    Frm56x.L9_Text = "BARCODE_CODE2"
ElseIf Frm56x.CBB7 = "Berat Riyal" Then
'ElseIf Frm56x.CBB7 = "Berat Amah" Then
    Frm56x.L9_Text = "BARCODE_RIYAL"
Else
    Frm56x.L9_Text = vbNullString
End If
End Sub
Private Sub CBB8_Click()
'On Error Resume Next
If Frm56x.CBB8 = "Berat" Then
    Frm56x.L10_Text = "BARCODE_BERAT"
ElseIf Frm56x.CBB8 = "Upah Modal" Then
    Frm56x.L10_Text = "BARCODE_UPAH"
ElseIf Frm56x.CBB8 = "Upah Jualan" Then
    Frm56x.L10_Text = "BARCODE_UPAH2"
ElseIf Frm56x.CBB8 = "Purity" Then
    Frm56x.L10_Text = "BARCODE_PURITY"
ElseIf Frm56x.CBB8 = "Panjang" Then
    Frm56x.L10_Text = "BARCODE_Panjang"
ElseIf Frm56x.CBB8 = "Lebar" Then
    Frm56x.L10_Text = "BARCODE_Lebar"
ElseIf Frm56x.CBB8 = "Saiz" Then
    Frm56x.L10_Text = "BARCODE_Saiz"
ElseIf Frm56x.CBB8 = "Dulang" Then
    Frm56x.L10_Text = "BARCODE_DULANG"
ElseIf Frm56x.CBB8 = "Supplier" Then
    Frm56x.L10_Text = "BARCODE_SUPPLIER"
ElseIf Frm56x.CBB8 = "Code 1" Then
    Frm56x.L10_Text = "BARCODE_CODE1"
ElseIf Frm56x.CBB8 = "Code 2" Then
    Frm56x.L10_Text = "BARCODE_CODE2"
ElseIf Frm56x.CBB8 = "Berat Riyal" Then
'ElseIf Frm56x.CBB8 = "Berat Amah" Then
    Frm56x.L10_Text = "BARCODE_RIYAL"
Else
    Frm56x.L10_Text = vbNullString
End If
End Sub
Private Sub CBB9_Click()
'On Error Resume Next
If Frm56x.CBB9 = "Berat" Then
    Frm56x.L11_Text = "BARCODE_BERAT"
ElseIf Frm56x.CBB9 = "Upah Modal" Then
    Frm56x.L11_Text = "BARCODE_UPAH"
ElseIf Frm56x.CBB9 = "Upah Jualan" Then
    Frm56x.L11_Text = "BARCODE_UPAH2"
ElseIf Frm56x.CBB9 = "Purity" Then
    Frm56x.L11_Text = "BARCODE_PURITY"
ElseIf Frm56x.CBB9 = "Panjang" Then
    Frm56x.L11_Text = "BARCODE_Panjang"
ElseIf Frm56x.CBB9 = "Lebar" Then
    Frm56x.L11_Text = "BARCODE_Lebar"
ElseIf Frm56x.CBB9 = "Saiz" Then
    Frm56x.L11_Text = "BARCODE_Saiz"
ElseIf Frm56x.CBB9 = "Dulang" Then
    Frm56x.L11_Text = "BARCODE_DULANG"
ElseIf Frm56x.CBB9 = "Supplier" Then
    Frm56x.L11_Text = "BARCODE_SUPPLIER"
ElseIf Frm56x.CBB9 = "Code 1" Then
    Frm56x.L11_Text = "BARCODE_CODE1"
ElseIf Frm56x.CBB9 = "Code 2" Then
    Frm56x.L11_Text = "BARCODE_CODE2"
ElseIf Frm56x.CBB9 = "Berat Riyal" Then
'ElseIf Frm56x.CBB9 = "Berat Amah" Then
    Frm56x.L11_Text = "BARCODE_RIYAL"
Else
    Frm56x.L11_Text = vbNullString
End If
End Sub
Private Sub CBB14_Click()
'On Error Resume Next
If Frm56x.CBB14 = "Berat" Then
    Frm56x.L12_Text = "BARCODE_BERAT"
ElseIf Frm56x.CBB14 = "Upah Modal" Then
    Frm56x.L12_Text = "BARCODE_UPAH"
ElseIf Frm56x.CBB14 = "Upah Jualan" Then
    Frm56x.L12_Text = "BARCODE_UPAH2"
ElseIf Frm56x.CBB14 = "Purity" Then
    Frm56x.L12_Text = "BARCODE_PURITY"
ElseIf Frm56x.CBB14 = "Panjang" Then
    Frm56x.L12_Text = "BARCODE_Panjang"
ElseIf Frm56x.CBB14 = "Lebar" Then
    Frm56x.L12_Text = "BARCODE_Lebar"
ElseIf Frm56x.CBB14 = "Saiz" Then
    Frm56x.L12_Text = "BARCODE_Saiz"
ElseIf Frm56x.CBB14 = "Dulang" Then
    Frm56x.L12_Text = "BARCODE_DULANG"
ElseIf Frm56x.CBB14 = "Supplier" Then
    Frm56x.L12_Text = "BARCODE_SUPPLIER"
ElseIf Frm56x.CBB14 = "Code 1" Then
    Frm56x.L12_Text = "BARCODE_CODE1"
ElseIf Frm56x.CBB14 = "Code 2" Then
    Frm56x.L12_Text = "BARCODE_CODE2"
ElseIf Frm56x.CBB14 = "Berat Riyal" Then
'ElseIf Frm56x.CBB14 = "Berat Amah" Then
    Frm56x.L12_Text = "BARCODE_RIYAL"
Else
    Frm56x.L12_Text = vbNullString
End If
End Sub
Private Sub CBB15_Click()
'On Error Resume Next
If Frm56x.CBB15 = "Berat" Then
    Frm56x.L13_Text = "BARCODE_BERAT"
ElseIf Frm56x.CBB15 = "Upah Modal" Then
    Frm56x.L13_Text = "BARCODE_UPAH"
ElseIf Frm56x.CBB15 = "Upah Jualan" Then
    Frm56x.L13_Text = "BARCODE_UPAH2"
ElseIf Frm56x.CBB15 = "Purity" Then
    Frm56x.L13_Text = "BARCODE_PURITY"
ElseIf Frm56x.CBB15 = "Panjang" Then
    Frm56x.L13_Text = "BARCODE_Panjang"
ElseIf Frm56x.CBB15 = "Lebar" Then
    Frm56x.L13_Text = "BARCODE_Lebar"
ElseIf Frm56x.CBB15 = "Saiz" Then
    Frm56x.L13_Text = "BARCODE_Saiz"
ElseIf Frm56x.CBB15 = "Dulang" Then
    Frm56x.L13_Text = "BARCODE_DULANG"
ElseIf Frm56x.CBB15 = "Supplier" Then
    Frm56x.L13_Text = "BARCODE_SUPPLIER"
ElseIf Frm56x.CBB15 = "Code 1" Then
    Frm56x.L13_Text = "BARCODE_CODE1"
ElseIf Frm56x.CBB15 = "Code 2" Then
    Frm56x.L13_Text = "BARCODE_CODE2"
ElseIf Frm56x.CBB15 = "Berat Riyal" Then
'ElseIf Frm56x.CBB15 = "Berat Amah" Then
    Frm56x.L13_Text = "BARCODE_RIYAL"
Else
    Frm56x.L13_Text = vbNullString
End If
End Sub
Private Sub CBB16_Click()
'On Error Resume Next
If Frm56x.CBB16 = "Berat" Then
    Frm56x.L14_Text = "BARCODE_BERAT"
ElseIf Frm56x.CBB16 = "Upah Modal" Then
    Frm56x.L14_Text = "BARCODE_UPAH"
ElseIf Frm56x.CBB16 = "Upah Jualan" Then
    Frm56x.L14_Text = "BARCODE_UPAH2"
ElseIf Frm56x.CBB16 = "Purity" Then
    Frm56x.L14_Text = "BARCODE_PURITY"
ElseIf Frm56x.CBB16 = "Panjang" Then
    Frm56x.L14_Text = "BARCODE_Panjang"
ElseIf Frm56x.CBB16 = "Lebar" Then
    Frm56x.L14_Text = "BARCODE_Lebar"
ElseIf Frm56x.CBB16 = "Saiz" Then
    Frm56x.L14_Text = "BARCODE_Saiz"
ElseIf Frm56x.CBB16 = "Dulang" Then
    Frm56x.L14_Text = "BARCODE_DULANG"
ElseIf Frm56x.CBB16 = "Supplier" Then
    Frm56x.L14_Text = "BARCODE_SUPPLIER"
ElseIf Frm56x.CBB16 = "Code 1" Then
    Frm56x.L14_Text = "BARCODE_CODE1"
ElseIf Frm56x.CBB16 = "Code 2" Then
    Frm56x.L14_Text = "BARCODE_CODE2"
ElseIf Frm56x.CBB16 = "Berat Riyal" Then
'ElseIf Frm56x.CBB16 = "Berat Amah" Then
    Frm56x.L14_Text = "BARCODE_RIYAL"
Else
    Frm56x.L14_Text = vbNullString
End If
End Sub
Private Sub CMD1_Click()
'On Error Resume Next
Dim Err(5)

If Frm56x.CBB10 = vbNullString Then
    x = x + 1
    Err(x) = "Sila pilih [Saiz tulisan barisan pertama]"
End If
If Frm56x.CBB11 = vbNullString Then
    x = x + 1
    Err(x) = "Sila pilih [Saiz tulisan barisan kedua]"
End If
If Frm56x.CBB12 = vbNullString Then
    x = x + 1
    Err(x) = "Sila pilih [Saiz tulisan barisan ketiga]"
End If
If Frm56x.CBB13 = vbNullString Then
    x = x + 1
    Err(x) = "Sila pilih [Saiz tulisan barisan keempat]"
End If
If Frm56x.CB14 = 0 And Frm56x.CB15 = 0 Then
    x = x + 1
    Err(x) = "Sila pilih jenis barcode label."
End If

If x <> 0 Then
    Frm6.Show
    Frm6.Pic1.Cls
    For Y = 1 To x
        Frm6.Pic1.Print Y & " - " & Err(Y)
    Next Y
    Exit Sub
Else

    Note = "Simpan tetapan ini ?" & vbCrLf & _
            "Sistem mungkin mengambil sedikit masa untuk menyimpan tetapan ini." & vbCrLf & _
            vbNullString & vbCrLf & _
            "Teruskan?"
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbYes Then
        Dim Frm56x_LM_BARCODE(12)
        Dim Frm56x_LM_Jenis(12)
        
        If MDI_frm1.L20_Text = "Semua cawangan" Then
            
            LM_KEDAI = "HQ"
            
        Else
            
            LM_KEDAI = MDI_frm1.L20_Text
        
        End If

        '#########Layout Barcode###########
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from layout_barcode where perkara='" & LM_KEDAI & "'", cn, adOpenKeyset, adLockOptimistic
    
        If Not rs.EOF Then
            If Frm56x.CB1 = 1 Then
                rs!Line1 = 1
            Else
                rs!Line1 = 0
            End If
            If Frm56x.CB2 = 1 Then
                rs!Line2 = 1
            Else
                rs!Line2 = 0
            End If
            If Frm56x.CB3 = 1 Then
                rs!Line3 = 1
            Else
                rs!Line3 = 0
            End If
            If Frm56x.CB4 = 1 Then
                rs!Line4 = 1
            Else
                rs!Line4 = 0
            End If
            
            rs!font_size_1 = Frm56x.CBB10 'Saiz tulisan bagi barisan pertama
            rs!font_size_2 = Frm56x.CBB11 'Saiz tulisan bagi barisan kedua
            rs!font_size_3 = Frm56x.CBB12 'Saiz tulisan bagi barisan ketiga
            rs!font_size_4 = Frm56x.CBB13 'Saiz tulisan bagi barisan keempat
            
            If Frm56x.CB14 = 1 Then '0 : Type A , 1 : Type B
                rs!BARCODE_TYPE = 0
            ElseIf Frm56x.CB15 = 1 Then
                rs!BARCODE_TYPE = 1
            End If
            
            rs.Update
        End If
        
        rs.Close
        Set rs = Nothing
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from tetapan_barcode where cawangan='" & LM_KEDAI & "'", cn, adOpenKeyset, adLockOptimistic
        
        While rs.EOF = False
            rs.Delete
            rs.Update
            rs.MoveNext
        Wend
        
        rs.Close
        Set rs = Nothing
        
        Frm56x_LM_Jenis(1) = Frm56x.CBB1
        Frm56x_LM_Jenis(2) = Frm56x.CBB2
        Frm56x_LM_Jenis(3) = Frm56x.CBB3
        Frm56x_LM_Jenis(4) = Frm56x.CBB4
        Frm56x_LM_Jenis(5) = Frm56x.CBB5
        Frm56x_LM_Jenis(6) = Frm56x.CBB6
        Frm56x_LM_Jenis(7) = Frm56x.CBB7
        Frm56x_LM_Jenis(8) = Frm56x.CBB8
        Frm56x_LM_Jenis(9) = Frm56x.CBB9
        Frm56x_LM_Jenis(10) = Frm56x.CBB14
        Frm56x_LM_Jenis(11) = Frm56x.CBB15
        Frm56x_LM_Jenis(12) = Frm56x.CBB16
        
        Frm56x_LM_BARCODE(1) = Frm56x.L3_Text
        Frm56x_LM_BARCODE(2) = Frm56x.L4_Text
        Frm56x_LM_BARCODE(3) = Frm56x.L5_Text
        Frm56x_LM_BARCODE(4) = Frm56x.L6_Text
        Frm56x_LM_BARCODE(5) = Frm56x.L7_Text
        Frm56x_LM_BARCODE(6) = Frm56x.L8_Text
        Frm56x_LM_BARCODE(7) = Frm56x.L9_Text
        Frm56x_LM_BARCODE(8) = Frm56x.L10_Text
        Frm56x_LM_BARCODE(9) = Frm56x.L11_Text
        Frm56x_LM_BARCODE(10) = Frm56x.L12_Text
        Frm56x_LM_BARCODE(11) = Frm56x.L13_Text
        Frm56x_LM_BARCODE(12) = Frm56x.L14_Text
        
        For i = 1 To 12
            If Frm56x_LM_BARCODE(i) = vbNullString Then
                Frm56x_LM_Jenis(i) = vbNullString
                Frm56x_LM_BARCODE(i) = "No Data"
            End If
        Next i
    
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from tetapan_barcode", cn, adOpenKeyset, adLockOptimistic
        
        For x = 1 To 12
            rs.AddNew
            rs!jenis = Frm56x_LM_Jenis(x)
            rs!Nama = Frm56x_LM_BARCODE(x)
            rs!cawangan = LM_KEDAI
            rs.Update
        Next x
        
        rs.Close
        Set rs = Nothing
        
        Call setting_barcode
        
        MsgBox "Tetapan BARCODE telah berjaya disimpan.", vbInformation, "Info"
    End If
End If
End Sub
Private Sub Form_Load()
'On Error Resume Next
Call Frm56x_ListItem
'Call Frm56x_Setting
End Sub
Private Sub Tmr1_Timer()
'On Error Resume Next
Frm56x.L1_Text = DateTime.Date
Frm56x.L2_Text = DateTime.Time$
End Sub
Private Sub Frm56x_ListItem()
'On Error Resume Next
With Frm56x.CBB1
    .AddItem vbNullString
    .AddItem "Berat"
    .AddItem "Upah Modal"
    .AddItem "Upah Jualan"
    .AddItem "Purity"
    .AddItem "Panjang"
    .AddItem "Lebar"
    .AddItem "Saiz"
    .AddItem "Dulang"
    .AddItem "Supplier"
    .AddItem "Code 1"
    .AddItem "Code 2"
    .AddItem "Berat Riyal"
End With

With Frm56x.CBB2
    .AddItem vbNullString
    .AddItem "Berat"
    .AddItem "Upah Modal"
    .AddItem "Upah Jualan"
    .AddItem "Purity"
    .AddItem "Panjang"
    .AddItem "Lebar"
    .AddItem "Saiz"
    .AddItem "Dulang"
    .AddItem "Supplier"
    .AddItem "Code 1"
    .AddItem "Code 2"
    .AddItem "Berat Riyal"
End With

With Frm56x.CBB3
    .AddItem vbNullString
    .AddItem "Berat"
    .AddItem "Upah Modal"
    .AddItem "Upah Jualan"
    .AddItem "Purity"
    .AddItem "Panjang"
    .AddItem "Lebar"
    .AddItem "Saiz"
    .AddItem "Dulang"
    .AddItem "Supplier"
    .AddItem "Code 1"
    .AddItem "Code 2"
    .AddItem "Berat Riyal"
End With

With Frm56x.CBB4
    .AddItem vbNullString
    .AddItem "Berat"
    .AddItem "Upah Modal"
    .AddItem "Upah Jualan"
    .AddItem "Purity"
    .AddItem "Panjang"
    .AddItem "Lebar"
    .AddItem "Saiz"
    .AddItem "Dulang"
    .AddItem "Supplier"
    .AddItem "Code 1"
    .AddItem "Code 2"
    .AddItem "Berat Riyal"
End With

With Frm56x.CBB5
    .AddItem vbNullString
    .AddItem "Berat"
    .AddItem "Upah Modal"
    .AddItem "Upah Jualan"
    .AddItem "Purity"
    .AddItem "Panjang"
    .AddItem "Lebar"
    .AddItem "Saiz"
    .AddItem "Dulang"
    .AddItem "Supplier"
    .AddItem "Code 1"
    .AddItem "Code 2"
    .AddItem "Berat Riyal"
End With

With Frm56x.CBB6
    .AddItem vbNullString
    .AddItem "Berat"
    .AddItem "Upah Modal"
    .AddItem "Upah Jualan"
    .AddItem "Purity"
    .AddItem "Panjang"
    .AddItem "Lebar"
    .AddItem "Saiz"
    .AddItem "Dulang"
    .AddItem "Supplier"
    .AddItem "Code 1"
    .AddItem "Code 2"
    .AddItem "Berat Riyal"
End With

With Frm56x.CBB7
    .AddItem vbNullString
    .AddItem "Berat"
    .AddItem "Upah Modal"
    .AddItem "Upah Jualan"
    .AddItem "Purity"
    .AddItem "Panjang"
    .AddItem "Lebar"
    .AddItem "Saiz"
    .AddItem "Dulang"
    .AddItem "Supplier"
    .AddItem "Code 1"
    .AddItem "Code 2"
    .AddItem "Berat Riyal"
End With

With Frm56x.CBB8
    .AddItem vbNullString
    .AddItem "Berat"
    .AddItem "Upah Modal"
    .AddItem "Upah Jualan"
    .AddItem "Purity"
    .AddItem "Panjang"
    .AddItem "Lebar"
    .AddItem "Saiz"
    .AddItem "Dulang"
    .AddItem "Supplier"
    .AddItem "Code 1"
    .AddItem "Code 2"
    .AddItem "Berat Riyal"
End With

With Frm56x.CBB9
    .AddItem vbNullString
    .AddItem "Berat"
    .AddItem "Upah Modal"
    .AddItem "Upah Jualan"
    .AddItem "Purity"
    .AddItem "Panjang"
    .AddItem "Lebar"
    .AddItem "Saiz"
    .AddItem "Dulang"
    .AddItem "Supplier"
    .AddItem "Code 1"
    .AddItem "Code 2"
    .AddItem "Berat Riyal"
End With

With Frm56x.CBB14
    .AddItem vbNullString
    .AddItem "Berat"
    .AddItem "Upah Modal"
    .AddItem "Upah Jualan"
    .AddItem "Purity"
    .AddItem "Panjang"
    .AddItem "Lebar"
    .AddItem "Saiz"
    .AddItem "Dulang"
    .AddItem "Supplier"
    .AddItem "Code 1"
    .AddItem "Code 2"
    .AddItem "Berat Riyal"
End With

With Frm56x.CBB15
    .AddItem vbNullString
    .AddItem "Berat"
    .AddItem "Upah Modal"
    .AddItem "Upah Jualan"
    .AddItem "Purity"
    .AddItem "Panjang"
    .AddItem "Lebar"
    .AddItem "Saiz"
    .AddItem "Dulang"
    .AddItem "Supplier"
    .AddItem "Code 1"
    .AddItem "Code 2"
    .AddItem "Berat Riyal"
End With

With Frm56x.CBB16
    .AddItem vbNullString
    .AddItem "Berat"
    .AddItem "Upah Modal"
    .AddItem "Upah Jualan"
    .AddItem "Purity"
    .AddItem "Panjang"
    .AddItem "Lebar"
    .AddItem "Saiz"
    .AddItem "Dulang"
    .AddItem "Supplier"
    .AddItem "Code 1"
    .AddItem "Code 2"
    .AddItem "Berat Riyal"
End With

Dim Frm56x_LM_BARCODE(12)
Dim Frm56x_LM_Jenis(12)

Frm56x.CBB10.Clear
Frm56x.CBB11.Clear
Frm56x.CBB12.Clear
Frm56x.CBB13.Clear

For i = 6 To 8

    Frm56x.CBB10.AddItem i
    Frm56x.CBB11.AddItem i
    Frm56x.CBB12.AddItem i
    Frm56x.CBB13.AddItem i
    
Next i

Frm56x.CBB10 = 6
Frm56x.CBB11 = 6
Frm56x.CBB12 = 6
Frm56x.CBB13 = 6
 
If MDI_frm1.L20_Text = "Semua cawangan" Then
    
    LM_KEDAI = "HQ"
    
Else
    
    LM_KEDAI = MDI_frm1.L20_Text

End If

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from layout_barcode where perkara='" & LM_KEDAI & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If rs!Line1 = 1 Then
        Frm56x.CB1 = 1
    Else
        Frm56x.CB1 = 0
    End If
    If rs!Line2 = 1 Then
        Frm56x.CB2 = 1
    Else
        Frm56x.CB2 = 0
    End If
    If rs!Line3 = 1 Then
        Frm56x.CB3 = 1
    Else
        Frm56x.CB3 = 0
    End If
    If rs!Line4 = 1 Then
        Frm56x.CB4 = 1
    Else
        Frm56x.CB4 = 0
    End If
    
    If Not IsNull(rs!BARCODE_TYPE) Then '0 : Type A , 1 : Type B
        
        If rs!BARCODE_TYPE = 0 Then
            Frm56x.CB14 = 1
            Frm56x.CB15 = 0
        ElseIf rs!BARCODE_TYPE = 1 Then
            Frm56x.CB15 = 1
            Frm56x.CB14 = 0
        End If
    
    End If

    On Error GoTo Err_A:
    If Not IsNull(rs!font_size_1) Then 'Saiz tulisan bagi barisan pertama
        Frm56x_LM_FONT_1 = rs!font_size_1
        Frm56x.CBB10 = Frm56x_LM_FONT_1
    End If
    
Restore_A:

    On Error GoTo Err_B:
    If Not IsNull(rs!font_size_2) Then 'Saiz tulisan bagi barisan kedua
        Frm56x_LM_FONT_2 = rs!font_size_2
        Frm56x.CBB11 = Frm56x_LM_FONT_2
    End If
    
Restore_B:

    On Error GoTo Err_C:
    If Not IsNull(rs!font_size_3) Then 'Saiz tulisan bagi barisan ketiga
        Frm56x_LM_FONT_3 = rs!font_size_3
        Frm56x.CBB12 = Frm56x_LM_FONT_3
    End If
    
Restore_C:

    On Error GoTo Err_D:
    If Not IsNull(rs!font_size_4) Then 'Saiz tulisan bagi barisan keempat
        Frm56x_LM_FONT_4 = rs!font_size_4
        Frm56x.CBB13 = Frm56x_LM_FONT_4
    End If
    
Restore_D:
    
    'on error resume next
End If

rs.Close
Set rs = Nothing

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from tetapan_barcode where cawangan='" & LM_KEDAI & "'", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    If Not IsNull(rs!jenis) Then Frm56x_LM_BARCODE(x) = rs!jenis
    If Not IsNull(rs!Nama) Then Frm56x_LM_Jenis(x) = rs!Nama
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

If Frm56x_LM_BARCODE(1) <> vbNullString Then Frm56x.CBB1 = Frm56x_LM_BARCODE(1)
If Frm56x_LM_BARCODE(2) <> vbNullString Then Frm56x.CBB2 = Frm56x_LM_BARCODE(2)
If Frm56x_LM_BARCODE(3) <> vbNullString Then Frm56x.CBB3 = Frm56x_LM_BARCODE(3)
If Frm56x_LM_BARCODE(4) <> vbNullString Then Frm56x.CBB4 = Frm56x_LM_BARCODE(4)
If Frm56x_LM_BARCODE(5) <> vbNullString Then Frm56x.CBB5 = Frm56x_LM_BARCODE(5)
If Frm56x_LM_BARCODE(6) <> vbNullString Then Frm56x.CBB6 = Frm56x_LM_BARCODE(6)
If Frm56x_LM_BARCODE(7) <> vbNullString Then Frm56x.CBB7 = Frm56x_LM_BARCODE(7)
If Frm56x_LM_BARCODE(8) <> vbNullString Then Frm56x.CBB8 = Frm56x_LM_BARCODE(8)
If Frm56x_LM_BARCODE(9) <> vbNullString Then Frm56x.CBB9 = Frm56x_LM_BARCODE(9)
If Frm56x_LM_BARCODE(10) <> vbNullString Then Frm56x.CBB14 = Frm56x_LM_BARCODE(10)
If Frm56x_LM_BARCODE(11) <> vbNullString Then Frm56x.CBB15 = Frm56x_LM_BARCODE(11)
If Frm56x_LM_BARCODE(12) <> vbNullString Then Frm56x.CBB16 = Frm56x_LM_BARCODE(12)

Frm56x.L3_Text = Frm56x_LM_Jenis(1)
Frm56x.L4_Text = Frm56x_LM_Jenis(2)
Frm56x.L5_Text = Frm56x_LM_Jenis(3)
Frm56x.L6_Text = Frm56x_LM_Jenis(4)
Frm56x.L7_Text = Frm56x_LM_Jenis(5)
Frm56x.L8_Text = Frm56x_LM_Jenis(6)
Frm56x.L9_Text = Frm56x_LM_Jenis(7)
Frm56x.L10_Text = Frm56x_LM_Jenis(8)
Frm56x.L11_Text = Frm56x_LM_Jenis(9)
Frm56x.L12_Text = Frm56x_LM_Jenis(10)
Frm56x.L13_Text = Frm56x_LM_Jenis(11)
Frm56x.L14_Text = Frm56x_LM_Jenis(12)

Exit Sub
Err_A:
Frm56x.CBB10.AddItem Frm56x_LM_FONT_1
Frm56x.CBB10 = Frm56x_LM_FONT_1
Resume Restore_A:

Exit Sub
Err_B:
Frm56x.CBB11.AddItem Frm56x_LM_FONT_2
Frm56x.CBB11 = Frm56x_LM_FONT_2
Resume Restore_B:

Exit Sub
Err_C:
Frm56x.CBB12.AddItem Frm56x_LM_FONT_3
Frm56x.CBB12 = Frm56x_LM_FONT_3
Resume Restore_C:

Exit Sub
Err_D:
Frm56x.CBB13.AddItem Frm56x_LM_FONT_4
Frm56x.CBB13 = Frm56x_LM_FONT_4
Resume Restore_D:

End Sub
Private Sub Frm56x_Setting()
'On Error Resume Next
Dim Frm56x_LM_BARCODE(9)
Dim Frm56x_LM_Jenis(9)
    
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from layout_barcode where perkara='" & "Barcode" & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If rs!Line1 = 1 Then
        Frm56x.CB1 = 1
    Else
        Frm56x.CB1 = 0
    End If
    If rs!Line2 = 1 Then
        Frm56x.CB2 = 1
    Else
        Frm56x.CB2 = 0
    End If
    If rs!Line3 = 1 Then
        Frm56x.CB3 = 1
    Else
        Frm56x.CB3 = 0
    End If
End If

rs.Close
Set rs = Nothing

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from tetapan_barcode", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    If Not IsNull(rs!jenis) Then Frm56x_LM_BARCODE(x) = rs!jenis
    If Not IsNull(rs!Nama) Then Frm56x_LM_Jenis(x) = rs!Nama
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

Frm56x.CBB1 = Frm56x_LM_Jenis(1)
Frm56x.CBB2 = Frm56x_LM_Jenis(2)
Frm56x.CBB3 = Frm56x_LM_Jenis(3)
Frm56x.CBB4 = Frm56x_LM_Jenis(4)
Frm56x.CBB5 = Frm56x_LM_Jenis(5)
Frm56x.CBB6 = Frm56x_LM_Jenis(6)
Frm56x.CBB7 = Frm56x_LM_Jenis(7)
Frm56x.CBB8 = Frm56x_LM_Jenis(8)
Frm56x.CBB9 = Frm56x_LM_Jenis(9)
Frm56x.CBB14 = Frm56x_LM_Jenis(10)
Frm56x.CBB15 = Frm56x_LM_Jenis(11)
Frm56x.CBB16 = Frm56x_LM_Jenis(12)

Frm56x.L3_Text = Frm56x_LM_BARCODE(1)
Frm56x.L4_Text = Frm56x_LM_BARCODE(2)
Frm56x.L5_Text = Frm56x_LM_BARCODE(3)
Frm56x.L6_Text = Frm56x_LM_BARCODE(4)
Frm56x.L7_Text = Frm56x_LM_BARCODE(5)
Frm56x.L8_Text = Frm56x_LM_BARCODE(6)
Frm56x.L9_Text = Frm56x_LM_BARCODE(7)
Frm56x.L10_Text = Frm56x_LM_BARCODE(8)
Frm56x.L11_Text = Frm56x_LM_BARCODE(9)
Frm56x.L12_Text = Frm56x_LM_BARCODE(10)
Frm56x.L13_Text = Frm56x_LM_BARCODE(11)
Frm56x.L14_Text = Frm56x_LM_BARCODE(12)
End Sub
