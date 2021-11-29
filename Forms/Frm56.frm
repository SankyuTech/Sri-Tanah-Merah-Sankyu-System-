VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm56 
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
   Icon            =   "Frm56.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   12930
   ScaleWidth      =   23760
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ListView LV1 
      Height          =   11655
      Left            =   120
      TabIndex        =   238
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   20558
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
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
      Height          =   2220
      Left            =   3840
      TabIndex        =   72
      Top             =   480
      Visible         =   0   'False
      Width           =   7935
      Begin VB.CheckBox CB16 
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
         Left            =   5280
         TabIndex        =   237
         Top             =   645
         Width           =   200
      End
      Begin VB.CommandButton CMD3 
         Caption         =   "Simpan Data"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   3120
         MouseIcon       =   "Frm56.frx":0ECA
         MousePointer    =   99  'Custom
         Picture         =   "Frm56.frx":11D4
         Style           =   1  'Graphical
         TabIndex        =   77
         Top             =   960
         Width           =   1935
      End
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
         Left            =   240
         TabIndex        =   74
         Top             =   645
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
         Left            =   2715
         TabIndex        =   73
         Top             =   645
         Width           =   200
      End
      Begin VB.Label Label52 
         BackStyle       =   0  'Transparent
         Caption         =   "Jenis barcode label yang digunakan."
         Height          =   255
         Left            =   240
         TabIndex        =   76
         Top             =   240
         Width           =   5295
      End
      Begin VB.Label Label46 
         BackStyle       =   0  'Transparent
         Caption         =   "Type A (35mm X 25mm)      Type B (75mm X 35mm)       Type C (70mm X 35mm)"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   480
         TabIndex        =   75
         Top             =   600
         Width           =   8655
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Developer Mode"
      Height          =   9375
      Left            =   2640
      TabIndex        =   71
      Top             =   2880
      Visible         =   0   'False
      Width           =   15735
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFFF&
         Height          =   7095
         Left            =   120
         TabIndex        =   78
         Top             =   2160
         Visible         =   0   'False
         Width           =   15480
         Begin VB.Frame Frame4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Element 1"
            Height          =   1815
            Index           =   0
            Left            =   120
            TabIndex        =   223
            Top             =   240
            Width           =   3735
            Begin VB.TextBox TB1 
               BackColor       =   &H00FFFFFF&
               Height          =   360
               Index           =   0
               Left            =   1395
               TabIndex        =   229
               Text            =   "TB1"
               Top             =   240
               Width           =   1450
            End
            Begin VB.TextBox TB2 
               BackColor       =   &H00FFFFFF&
               Height          =   360
               Index           =   0
               Left            =   1395
               TabIndex        =   228
               Text            =   "TB2"
               Top             =   600
               Width           =   1450
            End
            Begin VB.CheckBox CB10 
               Caption         =   "Scanner Mode"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   0
               Left            =   2880
               TabIndex        =   227
               Top             =   240
               Width           =   200
            End
            Begin VB.CheckBox CB11 
               Caption         =   "Scanner Mode"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   0
               Left            =   2880
               TabIndex        =   226
               Top             =   525
               Width           =   200
            End
            Begin VB.TextBox TB3 
               BackColor       =   &H00FFFFFF&
               Height          =   360
               Index           =   0
               Left            =   1395
               TabIndex        =   225
               Text            =   "TB3"
               Top             =   960
               Width           =   1450
            End
            Begin VB.TextBox TB4 
               BackColor       =   &H00FFFFFF&
               Height          =   360
               Index           =   0
               Left            =   1395
               TabIndex        =   224
               Text            =   "TB4"
               Top             =   1320
               Width           =   1450
            End
            Begin VB.Label Label31 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Size * :"
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   0
               Left            =   120
               TabIndex        =   235
               Top             =   285
               Width           =   1200
            End
            Begin VB.Label Label32 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Font Type * :"
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   0
               Left            =   120
               TabIndex        =   234
               Top             =   645
               Width           =   1200
            End
            Begin VB.Label Label33 
               BackStyle       =   0  'Transparent
               Caption         =   "Bold"
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   0
               Left            =   3120
               TabIndex        =   233
               Top             =   195
               Width           =   1080
            End
            Begin VB.Label Label34 
               BackStyle       =   0  'Transparent
               Caption         =   "Italic"
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   0
               Left            =   3120
               TabIndex        =   232
               Top             =   480
               Width           =   1080
            End
            Begin VB.Label Label35 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Position X * :"
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   0
               Left            =   120
               TabIndex        =   231
               Top             =   1005
               Width           =   1200
            End
            Begin VB.Label Label36 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Position Y * :"
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   0
               Left            =   120
               TabIndex        =   230
               Top             =   1365
               Width           =   1200
            End
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Element 2"
            Height          =   1815
            Index           =   1
            Left            =   120
            TabIndex        =   210
            Top             =   2040
            Width           =   3735
            Begin VB.TextBox TB4 
               BackColor       =   &H00FFFFFF&
               Height          =   360
               Index           =   1
               Left            =   1395
               TabIndex        =   216
               Text            =   "TB4"
               Top             =   1320
               Width           =   1450
            End
            Begin VB.TextBox TB3 
               BackColor       =   &H00FFFFFF&
               Height          =   360
               Index           =   1
               Left            =   1395
               TabIndex        =   215
               Text            =   "TB3"
               Top             =   960
               Width           =   1450
            End
            Begin VB.CheckBox CB11 
               Caption         =   "Scanner Mode"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   1
               Left            =   2880
               TabIndex        =   214
               Top             =   525
               Width           =   200
            End
            Begin VB.CheckBox CB10 
               Caption         =   "Scanner Mode"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   1
               Left            =   2880
               TabIndex        =   213
               Top             =   240
               Width           =   200
            End
            Begin VB.TextBox TB2 
               BackColor       =   &H00FFFFFF&
               Height          =   360
               Index           =   1
               Left            =   1395
               TabIndex        =   212
               Text            =   "TB2"
               Top             =   600
               Width           =   1450
            End
            Begin VB.TextBox TB1 
               BackColor       =   &H00FFFFFF&
               Height          =   360
               Index           =   1
               Left            =   1395
               TabIndex        =   211
               Text            =   "TB1"
               Top             =   240
               Width           =   1450
            End
            Begin VB.Label Label36 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Position Y * :"
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   1
               Left            =   120
               TabIndex        =   222
               Top             =   1365
               Width           =   1200
            End
            Begin VB.Label Label35 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Position X * :"
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   1
               Left            =   120
               TabIndex        =   221
               Top             =   1005
               Width           =   1200
            End
            Begin VB.Label Label34 
               BackStyle       =   0  'Transparent
               Caption         =   "Italic"
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   1
               Left            =   3120
               TabIndex        =   220
               Top             =   480
               Width           =   1080
            End
            Begin VB.Label Label33 
               BackStyle       =   0  'Transparent
               Caption         =   "Bold"
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   1
               Left            =   3120
               TabIndex        =   219
               Top             =   195
               Width           =   1080
            End
            Begin VB.Label Label32 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Font Type * :"
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   1
               Left            =   120
               TabIndex        =   218
               Top             =   645
               Width           =   1200
            End
            Begin VB.Label Label31 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Size * :"
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   1
               Left            =   120
               TabIndex        =   217
               Top             =   285
               Width           =   1200
            End
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Element 3"
            Height          =   1815
            Index           =   2
            Left            =   120
            TabIndex        =   197
            Top             =   3840
            Width           =   3735
            Begin VB.TextBox TB4 
               BackColor       =   &H00FFFFFF&
               Height          =   360
               Index           =   2
               Left            =   1395
               TabIndex        =   203
               Text            =   "TB4"
               Top             =   1320
               Width           =   1450
            End
            Begin VB.TextBox TB3 
               BackColor       =   &H00FFFFFF&
               Height          =   360
               Index           =   2
               Left            =   1395
               TabIndex        =   202
               Text            =   "TB3"
               Top             =   960
               Width           =   1450
            End
            Begin VB.CheckBox CB11 
               Caption         =   "Scanner Mode"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   2
               Left            =   2880
               TabIndex        =   201
               Top             =   525
               Width           =   200
            End
            Begin VB.CheckBox CB10 
               Caption         =   "Scanner Mode"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   2
               Left            =   2880
               TabIndex        =   200
               Top             =   240
               Width           =   200
            End
            Begin VB.TextBox TB2 
               BackColor       =   &H00FFFFFF&
               Height          =   360
               Index           =   2
               Left            =   1395
               TabIndex        =   199
               Text            =   "TB2"
               Top             =   600
               Width           =   1450
            End
            Begin VB.TextBox TB1 
               BackColor       =   &H00FFFFFF&
               Height          =   360
               Index           =   2
               Left            =   1395
               TabIndex        =   198
               Text            =   "TB1"
               Top             =   240
               Width           =   1450
            End
            Begin VB.Label Label36 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Position Y * :"
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   2
               Left            =   120
               TabIndex        =   209
               Top             =   1365
               Width           =   1200
            End
            Begin VB.Label Label35 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Position X * :"
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   2
               Left            =   120
               TabIndex        =   208
               Top             =   1005
               Width           =   1200
            End
            Begin VB.Label Label34 
               BackStyle       =   0  'Transparent
               Caption         =   "Italic"
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   2
               Left            =   3120
               TabIndex        =   207
               Top             =   480
               Width           =   1080
            End
            Begin VB.Label Label33 
               BackStyle       =   0  'Transparent
               Caption         =   "Bold"
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   2
               Left            =   3120
               TabIndex        =   206
               Top             =   195
               Width           =   1080
            End
            Begin VB.Label Label32 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Font Type * :"
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   2
               Left            =   120
               TabIndex        =   205
               Top             =   645
               Width           =   1200
            End
            Begin VB.Label Label31 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Size * :"
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   2
               Left            =   120
               TabIndex        =   204
               Top             =   285
               Width           =   1200
            End
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Element 4"
            Height          =   1815
            Index           =   3
            Left            =   3960
            TabIndex        =   184
            Top             =   240
            Width           =   3735
            Begin VB.TextBox TB4 
               BackColor       =   &H00FFFFFF&
               Height          =   360
               Index           =   3
               Left            =   1395
               TabIndex        =   190
               Text            =   "TB4"
               Top             =   1320
               Width           =   1450
            End
            Begin VB.TextBox TB3 
               BackColor       =   &H00FFFFFF&
               Height          =   360
               Index           =   3
               Left            =   1395
               TabIndex        =   189
               Text            =   "TB3"
               Top             =   960
               Width           =   1450
            End
            Begin VB.CheckBox CB11 
               Caption         =   "Scanner Mode"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   3
               Left            =   2880
               TabIndex        =   188
               Top             =   525
               Width           =   200
            End
            Begin VB.CheckBox CB10 
               Caption         =   "Scanner Mode"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   3
               Left            =   2880
               TabIndex        =   187
               Top             =   240
               Width           =   200
            End
            Begin VB.TextBox TB2 
               BackColor       =   &H00FFFFFF&
               Height          =   360
               Index           =   3
               Left            =   1395
               TabIndex        =   186
               Text            =   "TB2"
               Top             =   600
               Width           =   1450
            End
            Begin VB.TextBox TB1 
               BackColor       =   &H00FFFFFF&
               Height          =   360
               Index           =   3
               Left            =   1395
               TabIndex        =   185
               Text            =   "TB1"
               Top             =   240
               Width           =   1450
            End
            Begin VB.Label Label36 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Position Y * :"
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   3
               Left            =   120
               TabIndex        =   196
               Top             =   1365
               Width           =   1200
            End
            Begin VB.Label Label35 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Position X * :"
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   3
               Left            =   120
               TabIndex        =   195
               Top             =   1005
               Width           =   1200
            End
            Begin VB.Label Label34 
               BackStyle       =   0  'Transparent
               Caption         =   "Italic"
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   3
               Left            =   3120
               TabIndex        =   194
               Top             =   480
               Width           =   1080
            End
            Begin VB.Label Label33 
               BackStyle       =   0  'Transparent
               Caption         =   "Bold"
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   3
               Left            =   3120
               TabIndex        =   193
               Top             =   195
               Width           =   1080
            End
            Begin VB.Label Label32 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Font Type * :"
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   3
               Left            =   120
               TabIndex        =   192
               Top             =   645
               Width           =   1200
            End
            Begin VB.Label Label31 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Size * :"
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   3
               Left            =   120
               TabIndex        =   191
               Top             =   285
               Width           =   1200
            End
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Element 5"
            Height          =   1815
            Index           =   4
            Left            =   3960
            TabIndex        =   171
            Top             =   2040
            Width           =   3735
            Begin VB.TextBox TB4 
               BackColor       =   &H00FFFFFF&
               Height          =   360
               Index           =   4
               Left            =   1395
               TabIndex        =   177
               Text            =   "TB4"
               Top             =   1320
               Width           =   1450
            End
            Begin VB.TextBox TB3 
               BackColor       =   &H00FFFFFF&
               Height          =   360
               Index           =   4
               Left            =   1395
               TabIndex        =   176
               Text            =   "TB3"
               Top             =   960
               Width           =   1450
            End
            Begin VB.CheckBox CB11 
               Caption         =   "Scanner Mode"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   4
               Left            =   2880
               TabIndex        =   175
               Top             =   525
               Width           =   200
            End
            Begin VB.CheckBox CB10 
               Caption         =   "Scanner Mode"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   4
               Left            =   2880
               TabIndex        =   174
               Top             =   240
               Width           =   200
            End
            Begin VB.TextBox TB2 
               BackColor       =   &H00FFFFFF&
               Height          =   360
               Index           =   4
               Left            =   1395
               TabIndex        =   173
               Text            =   "TB2"
               Top             =   600
               Width           =   1450
            End
            Begin VB.TextBox TB1 
               BackColor       =   &H00FFFFFF&
               Height          =   360
               Index           =   4
               Left            =   1395
               TabIndex        =   172
               Text            =   "TB1"
               Top             =   240
               Width           =   1450
            End
            Begin VB.Label Label36 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Position Y * :"
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   4
               Left            =   120
               TabIndex        =   183
               Top             =   1365
               Width           =   1200
            End
            Begin VB.Label Label35 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Position X * :"
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   4
               Left            =   120
               TabIndex        =   182
               Top             =   1005
               Width           =   1200
            End
            Begin VB.Label Label34 
               BackStyle       =   0  'Transparent
               Caption         =   "Italic"
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   4
               Left            =   3120
               TabIndex        =   181
               Top             =   480
               Width           =   1080
            End
            Begin VB.Label Label33 
               BackStyle       =   0  'Transparent
               Caption         =   "Bold"
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   4
               Left            =   3120
               TabIndex        =   180
               Top             =   195
               Width           =   1080
            End
            Begin VB.Label Label32 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Font Type * :"
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   4
               Left            =   120
               TabIndex        =   179
               Top             =   645
               Width           =   1200
            End
            Begin VB.Label Label31 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Size * :"
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   4
               Left            =   120
               TabIndex        =   178
               Top             =   285
               Width           =   1200
            End
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Element 6"
            Height          =   1815
            Index           =   5
            Left            =   3960
            TabIndex        =   158
            Top             =   3840
            Width           =   3735
            Begin VB.TextBox TB4 
               BackColor       =   &H00FFFFFF&
               Height          =   360
               Index           =   5
               Left            =   1395
               TabIndex        =   164
               Text            =   "TB4"
               Top             =   1320
               Width           =   1450
            End
            Begin VB.TextBox TB3 
               BackColor       =   &H00FFFFFF&
               Height          =   360
               Index           =   5
               Left            =   1395
               TabIndex        =   163
               Text            =   "TB3"
               Top             =   960
               Width           =   1450
            End
            Begin VB.CheckBox CB11 
               Caption         =   "Scanner Mode"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   5
               Left            =   2880
               TabIndex        =   162
               Top             =   525
               Width           =   200
            End
            Begin VB.CheckBox CB10 
               Caption         =   "Scanner Mode"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   5
               Left            =   2880
               TabIndex        =   161
               Top             =   240
               Width           =   200
            End
            Begin VB.TextBox TB2 
               BackColor       =   &H00FFFFFF&
               Height          =   360
               Index           =   5
               Left            =   1395
               TabIndex        =   160
               Text            =   "TB2"
               Top             =   600
               Width           =   1450
            End
            Begin VB.TextBox TB1 
               BackColor       =   &H00FFFFFF&
               Height          =   360
               Index           =   5
               Left            =   1395
               TabIndex        =   159
               Text            =   "TB1"
               Top             =   240
               Width           =   1450
            End
            Begin VB.Label Label36 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Position Y * :"
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   5
               Left            =   120
               TabIndex        =   170
               Top             =   1365
               Width           =   1200
            End
            Begin VB.Label Label35 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Position X * :"
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   5
               Left            =   120
               TabIndex        =   169
               Top             =   1005
               Width           =   1200
            End
            Begin VB.Label Label34 
               BackStyle       =   0  'Transparent
               Caption         =   "Italic"
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   5
               Left            =   3120
               TabIndex        =   168
               Top             =   480
               Width           =   1080
            End
            Begin VB.Label Label33 
               BackStyle       =   0  'Transparent
               Caption         =   "Bold"
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   5
               Left            =   3120
               TabIndex        =   167
               Top             =   195
               Width           =   1080
            End
            Begin VB.Label Label32 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Font Type * :"
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   5
               Left            =   120
               TabIndex        =   166
               Top             =   645
               Width           =   1200
            End
            Begin VB.Label Label31 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Size * :"
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   5
               Left            =   120
               TabIndex        =   165
               Top             =   285
               Width           =   1200
            End
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Element 7"
            Height          =   1815
            Index           =   6
            Left            =   7800
            TabIndex        =   145
            Top             =   240
            Width           =   3735
            Begin VB.TextBox TB4 
               BackColor       =   &H00FFFFFF&
               Height          =   360
               Index           =   6
               Left            =   1395
               TabIndex        =   151
               Text            =   "TB4"
               Top             =   1320
               Width           =   1450
            End
            Begin VB.TextBox TB3 
               BackColor       =   &H00FFFFFF&
               Height          =   360
               Index           =   6
               Left            =   1395
               TabIndex        =   150
               Text            =   "TB3"
               Top             =   960
               Width           =   1450
            End
            Begin VB.CheckBox CB11 
               Caption         =   "Scanner Mode"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   6
               Left            =   2880
               TabIndex        =   149
               Top             =   525
               Width           =   200
            End
            Begin VB.CheckBox CB10 
               Caption         =   "Scanner Mode"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   6
               Left            =   2880
               TabIndex        =   148
               Top             =   240
               Width           =   200
            End
            Begin VB.TextBox TB2 
               BackColor       =   &H00FFFFFF&
               Height          =   360
               Index           =   6
               Left            =   1395
               TabIndex        =   147
               Text            =   "TB2"
               Top             =   600
               Width           =   1450
            End
            Begin VB.TextBox TB1 
               BackColor       =   &H00FFFFFF&
               Height          =   360
               Index           =   6
               Left            =   1395
               TabIndex        =   146
               Text            =   "TB1"
               Top             =   240
               Width           =   1450
            End
            Begin VB.Label Label36 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Position Y * :"
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   6
               Left            =   120
               TabIndex        =   157
               Top             =   1365
               Width           =   1200
            End
            Begin VB.Label Label35 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Position X * :"
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   6
               Left            =   120
               TabIndex        =   156
               Top             =   1005
               Width           =   1200
            End
            Begin VB.Label Label34 
               BackStyle       =   0  'Transparent
               Caption         =   "Italic"
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   6
               Left            =   3120
               TabIndex        =   155
               Top             =   480
               Width           =   1080
            End
            Begin VB.Label Label33 
               BackStyle       =   0  'Transparent
               Caption         =   "Bold"
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   6
               Left            =   3120
               TabIndex        =   154
               Top             =   195
               Width           =   1080
            End
            Begin VB.Label Label32 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Font Type * :"
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   6
               Left            =   120
               TabIndex        =   153
               Top             =   645
               Width           =   1200
            End
            Begin VB.Label Label31 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Size * :"
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   6
               Left            =   120
               TabIndex        =   152
               Top             =   285
               Width           =   1200
            End
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Element 8"
            Height          =   1815
            Index           =   7
            Left            =   7800
            TabIndex        =   132
            Top             =   2040
            Width           =   3735
            Begin VB.TextBox TB4 
               BackColor       =   &H00FFFFFF&
               Height          =   360
               Index           =   7
               Left            =   1395
               TabIndex        =   138
               Text            =   "TB4"
               Top             =   1320
               Width           =   1450
            End
            Begin VB.TextBox TB3 
               BackColor       =   &H00FFFFFF&
               Height          =   360
               Index           =   7
               Left            =   1395
               TabIndex        =   137
               Text            =   "TB3"
               Top             =   960
               Width           =   1450
            End
            Begin VB.CheckBox CB11 
               Caption         =   "Scanner Mode"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   7
               Left            =   2880
               TabIndex        =   136
               Top             =   525
               Width           =   200
            End
            Begin VB.CheckBox CB10 
               Caption         =   "Scanner Mode"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   7
               Left            =   2880
               TabIndex        =   135
               Top             =   240
               Width           =   200
            End
            Begin VB.TextBox TB2 
               BackColor       =   &H00FFFFFF&
               Height          =   360
               Index           =   7
               Left            =   1395
               TabIndex        =   134
               Text            =   "TB2"
               Top             =   600
               Width           =   1450
            End
            Begin VB.TextBox TB1 
               BackColor       =   &H00FFFFFF&
               Height          =   360
               Index           =   7
               Left            =   1395
               TabIndex        =   133
               Text            =   "TB1"
               Top             =   240
               Width           =   1450
            End
            Begin VB.Label Label36 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Position Y * :"
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   7
               Left            =   120
               TabIndex        =   144
               Top             =   1365
               Width           =   1200
            End
            Begin VB.Label Label35 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Position X * :"
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   7
               Left            =   120
               TabIndex        =   143
               Top             =   1005
               Width           =   1200
            End
            Begin VB.Label Label34 
               BackStyle       =   0  'Transparent
               Caption         =   "Italic"
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   7
               Left            =   3120
               TabIndex        =   142
               Top             =   480
               Width           =   1080
            End
            Begin VB.Label Label33 
               BackStyle       =   0  'Transparent
               Caption         =   "Bold"
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   7
               Left            =   3120
               TabIndex        =   141
               Top             =   195
               Width           =   1080
            End
            Begin VB.Label Label32 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Font Type * :"
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   7
               Left            =   120
               TabIndex        =   140
               Top             =   645
               Width           =   1200
            End
            Begin VB.Label Label31 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Size * :"
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   7
               Left            =   120
               TabIndex        =   139
               Top             =   285
               Width           =   1200
            End
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Element 9"
            Height          =   1815
            Index           =   8
            Left            =   7800
            TabIndex        =   119
            Top             =   3840
            Width           =   3735
            Begin VB.TextBox TB4 
               BackColor       =   &H00FFFFFF&
               Height          =   360
               Index           =   8
               Left            =   1395
               TabIndex        =   125
               Text            =   "TB4"
               Top             =   1320
               Width           =   1450
            End
            Begin VB.TextBox TB3 
               BackColor       =   &H00FFFFFF&
               Height          =   360
               Index           =   8
               Left            =   1395
               TabIndex        =   124
               Text            =   "TB3"
               Top             =   960
               Width           =   1450
            End
            Begin VB.CheckBox CB11 
               Caption         =   "Scanner Mode"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   8
               Left            =   2880
               TabIndex        =   123
               Top             =   525
               Width           =   200
            End
            Begin VB.CheckBox CB10 
               Caption         =   "Scanner Mode"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   8
               Left            =   2880
               TabIndex        =   122
               Top             =   240
               Width           =   200
            End
            Begin VB.TextBox TB2 
               BackColor       =   &H00FFFFFF&
               Height          =   360
               Index           =   8
               Left            =   1395
               TabIndex        =   121
               Text            =   "TB2"
               Top             =   600
               Width           =   1450
            End
            Begin VB.TextBox TB1 
               BackColor       =   &H00FFFFFF&
               Height          =   360
               Index           =   8
               Left            =   1395
               TabIndex        =   120
               Text            =   "TB1"
               Top             =   240
               Width           =   1450
            End
            Begin VB.Label Label36 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Position Y * :"
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   8
               Left            =   120
               TabIndex        =   131
               Top             =   1365
               Width           =   1200
            End
            Begin VB.Label Label35 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Position X * :"
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   8
               Left            =   120
               TabIndex        =   130
               Top             =   1005
               Width           =   1200
            End
            Begin VB.Label Label34 
               BackStyle       =   0  'Transparent
               Caption         =   "Italic"
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   8
               Left            =   3120
               TabIndex        =   129
               Top             =   480
               Width           =   1080
            End
            Begin VB.Label Label33 
               BackStyle       =   0  'Transparent
               Caption         =   "Bold"
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   8
               Left            =   3120
               TabIndex        =   128
               Top             =   195
               Width           =   1080
            End
            Begin VB.Label Label32 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Font Type * :"
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   8
               Left            =   120
               TabIndex        =   127
               Top             =   645
               Width           =   1200
            End
            Begin VB.Label Label31 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Size * :"
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   8
               Left            =   120
               TabIndex        =   126
               Top             =   285
               Width           =   1200
            End
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Element 10"
            Height          =   1815
            Index           =   9
            Left            =   11640
            TabIndex        =   106
            Top             =   240
            Width           =   3735
            Begin VB.TextBox TB4 
               BackColor       =   &H00FFFFFF&
               Height          =   360
               Index           =   9
               Left            =   1395
               TabIndex        =   112
               Text            =   "TB4"
               Top             =   1320
               Width           =   1450
            End
            Begin VB.TextBox TB3 
               BackColor       =   &H00FFFFFF&
               Height          =   360
               Index           =   9
               Left            =   1395
               TabIndex        =   111
               Text            =   "TB3"
               Top             =   960
               Width           =   1450
            End
            Begin VB.CheckBox CB11 
               Caption         =   "Scanner Mode"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   9
               Left            =   2880
               TabIndex        =   110
               Top             =   525
               Width           =   200
            End
            Begin VB.CheckBox CB10 
               Caption         =   "Scanner Mode"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   9
               Left            =   2880
               TabIndex        =   109
               Top             =   240
               Width           =   200
            End
            Begin VB.TextBox TB2 
               BackColor       =   &H00FFFFFF&
               Height          =   360
               Index           =   9
               Left            =   1395
               TabIndex        =   108
               Text            =   "TB2"
               Top             =   600
               Width           =   1450
            End
            Begin VB.TextBox TB1 
               BackColor       =   &H00FFFFFF&
               Height          =   360
               Index           =   9
               Left            =   1395
               TabIndex        =   107
               Text            =   "TB1"
               Top             =   240
               Width           =   1450
            End
            Begin VB.Label Label36 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Position Y * :"
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   9
               Left            =   120
               TabIndex        =   118
               Top             =   1365
               Width           =   1200
            End
            Begin VB.Label Label35 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Position X * :"
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   9
               Left            =   120
               TabIndex        =   117
               Top             =   1005
               Width           =   1200
            End
            Begin VB.Label Label34 
               BackStyle       =   0  'Transparent
               Caption         =   "Italic"
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   9
               Left            =   3120
               TabIndex        =   116
               Top             =   480
               Width           =   1080
            End
            Begin VB.Label Label33 
               BackStyle       =   0  'Transparent
               Caption         =   "Bold"
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   9
               Left            =   3120
               TabIndex        =   115
               Top             =   195
               Width           =   1080
            End
            Begin VB.Label Label32 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Font Type * :"
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   9
               Left            =   120
               TabIndex        =   114
               Top             =   645
               Width           =   1200
            End
            Begin VB.Label Label31 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Size * :"
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   9
               Left            =   120
               TabIndex        =   113
               Top             =   285
               Width           =   1200
            End
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Element 11"
            Height          =   1815
            Index           =   10
            Left            =   11640
            TabIndex        =   93
            Top             =   2040
            Width           =   3735
            Begin VB.TextBox TB4 
               BackColor       =   &H00FFFFFF&
               Height          =   360
               Index           =   10
               Left            =   1395
               TabIndex        =   99
               Text            =   "TB4"
               Top             =   1320
               Width           =   1450
            End
            Begin VB.TextBox TB3 
               BackColor       =   &H00FFFFFF&
               Height          =   360
               Index           =   10
               Left            =   1395
               TabIndex        =   98
               Text            =   "TB3"
               Top             =   960
               Width           =   1450
            End
            Begin VB.CheckBox CB11 
               Caption         =   "Scanner Mode"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   10
               Left            =   2880
               TabIndex        =   97
               Top             =   525
               Width           =   200
            End
            Begin VB.CheckBox CB10 
               Caption         =   "Scanner Mode"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   10
               Left            =   2880
               TabIndex        =   96
               Top             =   240
               Width           =   200
            End
            Begin VB.TextBox TB2 
               BackColor       =   &H00FFFFFF&
               Height          =   360
               Index           =   10
               Left            =   1395
               TabIndex        =   95
               Text            =   "TB2"
               Top             =   600
               Width           =   1450
            End
            Begin VB.TextBox TB1 
               BackColor       =   &H00FFFFFF&
               Height          =   360
               Index           =   10
               Left            =   1395
               TabIndex        =   94
               Text            =   "TB1"
               Top             =   240
               Width           =   1450
            End
            Begin VB.Label Label36 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Position Y * :"
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   10
               Left            =   120
               TabIndex        =   105
               Top             =   1365
               Width           =   1200
            End
            Begin VB.Label Label35 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Position X * :"
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   10
               Left            =   120
               TabIndex        =   104
               Top             =   1005
               Width           =   1200
            End
            Begin VB.Label Label34 
               BackStyle       =   0  'Transparent
               Caption         =   "Italic"
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   10
               Left            =   3120
               TabIndex        =   103
               Top             =   480
               Width           =   1080
            End
            Begin VB.Label Label33 
               BackStyle       =   0  'Transparent
               Caption         =   "Bold"
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   10
               Left            =   3120
               TabIndex        =   102
               Top             =   195
               Width           =   1080
            End
            Begin VB.Label Label32 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Font Type * :"
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   10
               Left            =   120
               TabIndex        =   101
               Top             =   645
               Width           =   1200
            End
            Begin VB.Label Label31 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Size * :"
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   10
               Left            =   120
               TabIndex        =   100
               Top             =   285
               Width           =   1200
            End
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Element 12"
            Height          =   1815
            Index           =   11
            Left            =   11640
            TabIndex        =   80
            Top             =   3840
            Width           =   3735
            Begin VB.TextBox TB4 
               BackColor       =   &H00FFFFFF&
               Height          =   360
               Index           =   11
               Left            =   1395
               TabIndex        =   86
               Text            =   "TB4"
               Top             =   1320
               Width           =   1450
            End
            Begin VB.TextBox TB3 
               BackColor       =   &H00FFFFFF&
               Height          =   360
               Index           =   11
               Left            =   1395
               TabIndex        =   85
               Text            =   "TB3"
               Top             =   960
               Width           =   1450
            End
            Begin VB.CheckBox CB11 
               Caption         =   "Scanner Mode"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   11
               Left            =   2880
               TabIndex        =   84
               Top             =   525
               Width           =   200
            End
            Begin VB.CheckBox CB10 
               Caption         =   "Scanner Mode"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   11
               Left            =   2880
               TabIndex        =   83
               Top             =   240
               Width           =   200
            End
            Begin VB.TextBox TB2 
               BackColor       =   &H00FFFFFF&
               Height          =   360
               Index           =   11
               Left            =   1395
               TabIndex        =   82
               Text            =   "TB2"
               Top             =   600
               Width           =   1450
            End
            Begin VB.TextBox TB1 
               BackColor       =   &H00FFFFFF&
               Height          =   360
               Index           =   11
               Left            =   1395
               TabIndex        =   81
               Text            =   "TB1"
               Top             =   240
               Width           =   1450
            End
            Begin VB.Label Label36 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Position Y * :"
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   11
               Left            =   120
               TabIndex        =   92
               Top             =   1365
               Width           =   1200
            End
            Begin VB.Label Label35 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Position X * :"
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   11
               Left            =   120
               TabIndex        =   91
               Top             =   1005
               Width           =   1200
            End
            Begin VB.Label Label34 
               BackStyle       =   0  'Transparent
               Caption         =   "Italic"
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   11
               Left            =   3120
               TabIndex        =   90
               Top             =   480
               Width           =   1080
            End
            Begin VB.Label Label33 
               BackStyle       =   0  'Transparent
               Caption         =   "Bold"
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   11
               Left            =   3120
               TabIndex        =   89
               Top             =   195
               Width           =   1080
            End
            Begin VB.Label Label32 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Font Type * :"
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   11
               Left            =   120
               TabIndex        =   88
               Top             =   645
               Width           =   1200
            End
            Begin VB.Label Label31 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Size * :"
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   11
               Left            =   120
               TabIndex        =   87
               Top             =   285
               Width           =   1200
            End
         End
         Begin VB.CommandButton CMD2 
            Caption         =   "Simpan Data"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1095
            Left            =   6480
            MouseIcon       =   "Frm56.frx":379E
            MousePointer    =   99  'Custom
            Picture         =   "Frm56.frx":3AA8
            Style           =   1  'Graphical
            TabIndex        =   79
            Top             =   5760
            Width           =   1815
         End
         Begin VB.Label L15_Text 
            BackColor       =   &H8000000A&
            Caption         =   "L15_Text"
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   1560
            TabIndex        =   236
            Top             =   6480
            Visible         =   0   'False
            Width           =   1080
         End
      End
      Begin MSComctlLib.ListView LV2 
         Height          =   1755
         Left            =   120
         TabIndex        =   239
         Top             =   360
         Width           =   15495
         _ExtentX        =   27331
         _ExtentY        =   3096
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   4935
      Left            =   1680
      TabIndex        =   21
      Top             =   120
      Visible         =   0   'False
      Width           =   18975
      Begin VB.CommandButton CMD1 
         Caption         =   "Simpan Data"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   4920
         MouseIcon       =   "Frm56.frx":6072
         MousePointer    =   99  'Custom
         Picture         =   "Frm56.frx":637C
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   2520
         Width           =   1935
      End
      Begin VB.Timer Tmr1 
         Interval        =   10
         Left            =   10920
         Top             =   2880
      End
      Begin VB.ComboBox CBB1 
         Height          =   360
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   45
         Top             =   630
         Width           =   2415
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
         TabIndex        =   44
         Top             =   720
         Width           =   200
      End
      Begin VB.ComboBox CBB2 
         Height          =   360
         Left            =   4680
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   630
         Width           =   2415
      End
      Begin VB.ComboBox CBB3 
         Height          =   360
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Top             =   1110
         Width           =   2415
      End
      Begin VB.ComboBox CBB4 
         Height          =   360
         Left            =   4680
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   1110
         Width           =   2415
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
         TabIndex        =   40
         Top             =   1200
         Width           =   200
      End
      Begin VB.ComboBox CBB5 
         Height          =   360
         Left            =   7440
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   1110
         Width           =   2415
      End
      Begin VB.ComboBox CBB6 
         Height          =   360
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   1590
         Width           =   2415
      End
      Begin VB.ComboBox CBB7 
         Height          =   360
         Left            =   4680
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   1560
         Width           =   2415
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
         TabIndex        =   36
         Top             =   1680
         Width           =   200
      End
      Begin VB.ComboBox CBB8 
         Height          =   360
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   2040
         Width           =   2415
      End
      Begin VB.ComboBox CBB9 
         Height          =   360
         Left            =   4680
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   2040
         Width           =   2415
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
         TabIndex        =   33
         Top             =   2115
         Width           =   200
      End
      Begin VB.ComboBox CBB10 
         Height          =   360
         Left            =   12840
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   630
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.ComboBox CBB11 
         Height          =   360
         Left            =   12840
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   1110
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.ComboBox CBB12 
         Height          =   360
         Left            =   12840
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   1560
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.ComboBox CBB13 
         Height          =   360
         Left            =   12840
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   2040
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.ComboBox CBB14 
         Height          =   360
         Left            =   7440
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   630
         Width           =   2415
      End
      Begin VB.ComboBox CBB15 
         Height          =   360
         Left            =   7440
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   1560
         Width           =   2415
      End
      Begin VB.ComboBox CBB16 
         Height          =   360
         Left            =   7440
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   2040
         Width           =   2415
      End
      Begin VB.ComboBox CBB20 
         Height          =   360
         Left            =   15120
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   600
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.ComboBox CBB21 
         Height          =   360
         Left            =   15120
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   1080
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.ComboBox CBB22 
         Height          =   360
         Left            =   15120
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   1560
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.ComboBox CBB23 
         Height          =   360
         Left            =   15120
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   2040
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Sila buat penetapan barcode yang akan dicetak pada setiap item."
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   120
         TabIndex        =   69
         Top             =   240
         Width           =   9240
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Barisan Pertama"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   360
         TabIndex        =   68
         Top             =   675
         Width           =   1560
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "/"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4320
         TabIndex        =   67
         Top             =   675
         Width           =   360
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "/"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7080
         TabIndex        =   66
         Top             =   1155
         Width           =   360
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Barisan Kedua"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   360
         TabIndex        =   65
         Top             =   1155
         Width           =   1680
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "/"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4320
         TabIndex        =   64
         Top             =   1155
         Width           =   360
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Barisan Ketiga "
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   360
         TabIndex        =   63
         Top             =   1635
         Width           =   1680
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "/"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4320
         TabIndex        =   62
         Top             =   1635
         Width           =   360
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "/"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4320
         TabIndex        =   61
         Top             =   2070
         Width           =   360
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Barisan Keempat"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   360
         TabIndex        =   60
         Top             =   2070
         Width           =   1680
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Saiz tulisan barisan pertama"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   9960
         TabIndex        =   59
         Top             =   675
         Visible         =   0   'False
         Width           =   3000
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Saiz tulisan barisan kedua"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   9960
         TabIndex        =   58
         Top             =   1155
         Visible         =   0   'False
         Width           =   3000
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Saiz tulisan barisan ketiga"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   9960
         TabIndex        =   57
         Top             =   1635
         Visible         =   0   'False
         Width           =   3000
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Saiz tulisan barisan keempat"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   9960
         TabIndex        =   56
         Top             =   2070
         Visible         =   0   'False
         Width           =   3000
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "/"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7080
         TabIndex        =   55
         Top             =   675
         Width           =   360
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "/"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7080
         TabIndex        =   54
         Top             =   1635
         Width           =   360
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "/"
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   0
         Left            =   7080
         TabIndex        =   53
         Top             =   2070
         Width           =   360
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "1) User perlu log out dari sistem (Semua station) bagi memboleh tetapan ini diupdate selepas tetapan ini disimpan."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   52
         Top             =   3720
         Width           =   12135
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "Font Type :"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   14040
         TabIndex        =   51
         Top             =   645
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "Font Type :"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   14040
         TabIndex        =   50
         Top             =   1125
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Font Type :"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   14040
         TabIndex        =   49
         Top             =   1605
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "Font Type :"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   14040
         TabIndex        =   48
         Top             =   2085
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "2) Anda perlu membuat pemilihan data yang akan dicetakkan di atas tag dan kemungkinan besar bukan semua maklumat"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   47
         Top             =   4080
         Width           =   12135
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "dapat dicetak di atas tag di atas faktor saiz tag."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   360
         TabIndex        =   46
         Top             =   4440
         Width           =   12135
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   3495
      Left            =   2760
      ScaleHeight     =   3495
      ScaleWidth      =   6015
      TabIndex        =   11
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
         TabIndex        =   17
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
         TabIndex        =   16
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
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   13
         Top             =   1680
         Width           =   5415
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Preview :"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Width           =   1680
      End
      Begin VB.Image Image1 
         Height          =   750
         Left            =   120
         Picture         =   "Frm56.frx":8946
         Top             =   480
         Width           =   2790
      End
   End
   Begin MSComctlLib.ImageList ImageList4 
      Left            =   5640
      Top             =   6720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   60
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm56.frx":95D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm56.frx":BBAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm56.frx":E184
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm56.frx":1075E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm56.frx":12D38
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm56.frx":15312
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm56.frx":178EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm56.frx":19EC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm56.frx":1C4A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm56.frx":1EA7A
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm56.frx":21054
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm56.frx":2362E
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm56.frx":25C08
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm56.frx":281E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm56.frx":2A7BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm56.frx":2CD96
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm56.frx":2F370
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm56.frx":3194A
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm56.frx":33F24
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm56.frx":364FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm56.frx":38AD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm56.frx":3B0B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm56.frx":3D68C
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm56.frx":3FC66
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm56.frx":42240
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm56.frx":4481A
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm56.frx":46DF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm56.frx":493CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm56.frx":4B9A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm56.frx":4DF82
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm56.frx":5055C
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm56.frx":52B36
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm56.frx":55110
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm56.frx":576EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm56.frx":59CC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm56.frx":5C29E
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm56.frx":5E878
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm56.frx":60E52
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm56.frx":6342C
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm56.frx":65A06
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm56.frx":67FE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm56.frx":6A5BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm56.frx":6CB94
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm56.frx":6F16E
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm56.frx":71748
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm56.frx":73D22
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm56.frx":762FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm56.frx":788D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm56.frx":7AEB0
            Key             =   ""
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm56.frx":7D48A
            Key             =   ""
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm56.frx":7FA64
            Key             =   ""
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm56.frx":8203E
            Key             =   ""
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm56.frx":AA452
            Key             =   ""
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm56.frx":ACA2C
            Key             =   ""
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm56.frx":AF006
            Key             =   ""
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm56.frx":B15E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm56.frx":B3BBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm56.frx":B6194
            Key             =   ""
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm56.frx":B876E
            Key             =   ""
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm56.frx":BAD48
            Key             =   ""
         EndProperty
      EndProperty
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
      TabIndex        =   20
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
      TabIndex        =   19
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
      TabIndex        =   18
      Top             =   8400
      Visible         =   0   'False
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
      TabIndex        =   10
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
      TabIndex        =   9
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
      TabIndex        =   8
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
      TabIndex        =   7
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
      TabIndex        =   6
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
      TabIndex        =   5
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
      TabIndex        =   4
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
      TabIndex        =   3
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
      TabIndex        =   2
      Top             =   6960
      Visible         =   0   'False
      Width           =   1680
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
      TabIndex        =   1
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
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   2100
   End
End
Attribute VB_Name = "Frm56"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CB14_Click()
'On Error GoTo logging:
If Frm56.CB14 = 1 Then
    Frm56.CB15 = 0
    Frm56.CB16 = 0
End If

Exit Sub

logging:

LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " Frm56 : CB14_Click" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
Call log_rekod

Resume Next
End Sub
Private Sub CB15_Click()
'On Error GoTo logging:
If Frm56.CB15 = 1 Then
    Frm56.CB14 = 0
    Frm56.CB16 = 0
End If

Exit Sub

logging:

LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " Frm56 : CB15_Click" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
Call log_rekod

Resume Next
End Sub
Private Sub CB16_Click()
'On Error GoTo logging:
If Frm56.CB16 = 1 Then
    Frm56.CB14 = 0
    Frm56.CB15 = 0
End If

Exit Sub

logging:

LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " Frm56 : CB16_Click" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
Call log_rekod

Resume Next
End Sub
Private Sub CBB1_Click()
'On Error GoTo logging:
If Frm56.CBB1 = "Berat" Then
    Frm56.L3_Text = "BARCODE_BERAT"
ElseIf Frm56.CBB1 = "Upah Modal" Then
    Frm56.L3_Text = "BARCODE_UPAH"
ElseIf Frm56.CBB1 = "Upah Jualan" Then
    Frm56.L3_Text = "BARCODE_UPAH2"
ElseIf Frm56.CBB1 = "Purity" Then
    Frm56.L3_Text = "BARCODE_PURITY"
ElseIf Frm56.CBB1 = "Panjang" Then
    Frm56.L3_Text = "BARCODE_Panjang"
ElseIf Frm56.CBB1 = "Lebar" Then
    Frm56.L3_Text = "BARCODE_Lebar"
ElseIf Frm56.CBB1 = "Saiz" Then
    Frm56.L3_Text = "BARCODE_Saiz"
ElseIf Frm56.CBB1 = "Dulang" Then
    Frm56.L3_Text = "BARCODE_DULANG"
ElseIf Frm56.CBB1 = "Supplier" Then
    Frm56.L3_Text = "BARCODE_SUPPLIER"
ElseIf Frm56.CBB1 = "Code 1" Then
    Frm56.L3_Text = "BARCODE_CODE1"
ElseIf Frm56.CBB1 = "Code 2" Then
    Frm56.L3_Text = "BARCODE_CODE2"
ElseIf Frm56.CBB1 = "Barcode" Then
    Frm56.L3_Text = "BARCODE_BARCODE"
ElseIf Frm56.CBB1 = "Berat Riyal" Then
    Frm56.L3_Text = "BARCODE_RIYAL"
ElseIf Frm56.CBB1 = "Harga" Then
    Frm56.L3_Text = "BARCODE_HARGA"
ElseIf Frm56.CBB1 = "Modal" Then
    Frm56.L3_Text = "BARCODE_MODAL"
ElseIf Frm56.CBB1 = "Diamond" Then
    Frm56.L3_Text = "BARCODE_DIAMOND"
ElseIf Frm56.CBB1 = "Design" Then
    Frm56.L3_Text = "BARCODE_DESIGN"
ElseIf Frm56.CBB1 = "No Data" Then
    Frm56.L3_Text = "No Data"
Else
    Frm56.L3_Text = vbNullString
End If

Exit Sub

logging:

LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " Frm56 : CBB1_Click" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
Call log_rekod

Resume Next
End Sub
Private Sub CBB2_Click()
'On Error GoTo logging:
If Frm56.CBB2 = "Berat" Then
    Frm56.L4_Text = "BARCODE_BERAT"
ElseIf Frm56.CBB2 = "Upah Modal" Then
    Frm56.L4_Text = "BARCODE_UPAH"
ElseIf Frm56.CBB2 = "Upah Jualan" Then
    Frm56.L4_Text = "BARCODE_UPAH2"
ElseIf Frm56.CBB2 = "Purity" Then
    Frm56.L4_Text = "BARCODE_PURITY"
ElseIf Frm56.CBB2 = "Panjang" Then
    Frm56.L4_Text = "BARCODE_Panjang"
ElseIf Frm56.CBB2 = "Lebar" Then
    Frm56.L4_Text = "BARCODE_Lebar"
ElseIf Frm56.CBB2 = "Saiz" Then
    Frm56.L4_Text = "BARCODE_Saiz"
ElseIf Frm56.CBB2 = "Dulang" Then
    Frm56.L4_Text = "BARCODE_DULANG"
ElseIf Frm56.CBB2 = "Supplier" Then
    Frm56.L4_Text = "BARCODE_SUPPLIER"
ElseIf Frm56.CBB2 = "Code 1" Then
    Frm56.L4_Text = "BARCODE_CODE1"
ElseIf Frm56.CBB2 = "Code 2" Then
    Frm56.L4_Text = "BARCODE_CODE2"
ElseIf Frm56.CBB2 = "Barcode" Then
    Frm56.L4_Text = "BARCODE_BARCODE"
ElseIf Frm56.CBB2 = "Berat Riyal" Then
    Frm56.L4_Text = "BARCODE_RIYAL"
ElseIf Frm56.CBB2 = "Harga" Then
    Frm56.L4_Text = "BARCODE_HARGA"
ElseIf Frm56.CBB2 = "Modal" Then
    Frm56.L4_Text = "BARCODE_MODAL"
ElseIf Frm56.CBB2 = "Diamond" Then
    Frm56.L4_Text = "BARCODE_DIAMOND"
ElseIf Frm56.CBB2 = "Design" Then
    Frm56.L4_Text = "BARCODE_DESIGN"
ElseIf Frm56.CBB2 = "No Data" Then
    Frm56.L4_Text = "No Data"
Else
    Frm56.L4_Text = vbNullString
End If

Exit Sub

logging:

LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " Frm56 : CBB2_Click" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
Call log_rekod

Resume Next
End Sub
Private Sub CBB3_Click()
'On Error GoTo logging:
If Frm56.CBB3 = "Berat" Then
    Frm56.L5_Text = "BARCODE_BERAT"
ElseIf Frm56.CBB3 = "Upah Modal" Then
    Frm56.L5_Text = "BARCODE_UPAH"
ElseIf Frm56.CBB3 = "Upah Jualan" Then
    Frm56.L5_Text = "BARCODE_UPAH2"
ElseIf Frm56.CBB3 = "Purity" Then
    Frm56.L5_Text = "BARCODE_PURITY"
ElseIf Frm56.CBB3 = "Panjang" Then
    Frm56.L5_Text = "BARCODE_Panjang"
ElseIf Frm56.CBB3 = "Lebar" Then
    Frm56.L5_Text = "BARCODE_Lebar"
ElseIf Frm56.CBB3 = "Saiz" Then
    Frm56.L5_Text = "BARCODE_Saiz"
ElseIf Frm56.CBB3 = "Dulang" Then
    Frm56.L5_Text = "BARCODE_DULANG"
ElseIf Frm56.CBB3 = "Supplier" Then
    Frm56.L5_Text = "BARCODE_SUPPLIER"
ElseIf Frm56.CBB3 = "Code 1" Then
    Frm56.L5_Text = "BARCODE_CODE1"
ElseIf Frm56.CBB3 = "Code 2" Then
    Frm56.L5_Text = "BARCODE_CODE2"
ElseIf Frm56.CBB3 = "Barcode" Then
    Frm56.L5_Text = "BARCODE_BARCODE"
ElseIf Frm56.CBB3 = "Berat Riyal" Then
    Frm56.L5_Text = "BARCODE_RIYAL"
ElseIf Frm56.CBB3 = "Harga" Then
    Frm56.L5_Text = "BARCODE_HARGA"
ElseIf Frm56.CBB3 = "Modal" Then
    Frm56.L5_Text = "BARCODE_MODAL"
ElseIf Frm56.CBB3 = "Diamond" Then
    Frm56.L5_Text = "BARCODE_DIAMOND"
ElseIf Frm56.CBB3 = "Design" Then
    Frm56.L5_Text = "BARCODE_DESIGN"
ElseIf Frm56.CBB3 = "No Data" Then
    Frm56.L5_Text = "No Data"
Else
    Frm56.L5_Text = vbNullString
End If

Exit Sub

logging:

LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " Frm56 : CBB3_Click" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
Call log_rekod

Resume Next
End Sub
Private Sub CBB4_Click()
'On Error GoTo logging:
If Frm56.CBB4 = "Berat" Then
    Frm56.L6_Text = "BARCODE_BERAT"
ElseIf Frm56.CBB4 = "Upah Modal" Then
    Frm56.L6_Text = "BARCODE_UPAH"
ElseIf Frm56.CBB4 = "Upah Jualan" Then
    Frm56.L6_Text = "BARCODE_UPAH2"
ElseIf Frm56.CBB4 = "Purity" Then
    Frm56.L6_Text = "BARCODE_PURITY"
ElseIf Frm56.CBB4 = "Panjang" Then
    Frm56.L6_Text = "BARCODE_Panjang"
ElseIf Frm56.CBB4 = "Lebar" Then
    Frm56.L6_Text = "BARCODE_Lebar"
ElseIf Frm56.CBB4 = "Saiz" Then
    Frm56.L6_Text = "BARCODE_Saiz"
ElseIf Frm56.CBB4 = "Dulang" Then
    Frm56.L6_Text = "BARCODE_DULANG"
ElseIf Frm56.CBB4 = "Supplier" Then
    Frm56.L6_Text = "BARCODE_SUPPLIER"
ElseIf Frm56.CBB4 = "Code 1" Then
    Frm56.L6_Text = "BARCODE_CODE1"
ElseIf Frm56.CBB4 = "Code 2" Then
    Frm56.L6_Text = "BARCODE_CODE2"
ElseIf Frm56.CBB4 = "Barcode" Then
    Frm56.L6_Text = "BARCODE_BARCODE"
ElseIf Frm56.CBB4 = "Berat Riyal" Then
    Frm56.L6_Text = "BARCODE_RIYAL"
ElseIf Frm56.CBB4 = "Harga" Then
    Frm56.L6_Text = "BARCODE_HARGA"
ElseIf Frm56.CBB4 = "Modal" Then
    Frm56.L6_Text = "BARCODE_MODAL"
ElseIf Frm56.CBB4 = "Diamond" Then
    Frm56.L6_Text = "BARCODE_DIAMOND"
ElseIf Frm56.CBB4 = "Design" Then
    Frm56.L6_Text = "BARCODE_DESIGN"
ElseIf Frm56.CBB4 = "No Data" Then
    Frm56.L6_Text = "No Data"
Else
    Frm56.L6_Text = vbNullString
End If

Exit Sub

logging:

LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " Frm56 : CBB4_Click" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
Call log_rekod

Resume Next
End Sub
Private Sub CBB5_Click()
'On Error GoTo logging:
If Frm56.CBB5 = "Berat" Then
    Frm56.L7_Text = "BARCODE_BERAT"
ElseIf Frm56.CBB5 = "Upah Modal" Then
    Frm56.L7_Text = "BARCODE_UPAH"
ElseIf Frm56.CBB5 = "Upah Jualan" Then
    Frm56.L7_Text = "BARCODE_UPAH2"
ElseIf Frm56.CBB5 = "Purity" Then
    Frm56.L7_Text = "BARCODE_PURITY"
ElseIf Frm56.CBB5 = "Panjang" Then
    Frm56.L7_Text = "BARCODE_Panjang"
ElseIf Frm56.CBB5 = "Lebar" Then
    Frm56.L7_Text = "BARCODE_Lebar"
ElseIf Frm56.CBB5 = "Saiz" Then
    Frm56.L7_Text = "BARCODE_Saiz"
ElseIf Frm56.CBB5 = "Dulang" Then
    Frm56.L7_Text = "BARCODE_DULANG"
ElseIf Frm56.CBB5 = "Supplier" Then
    Frm56.L7_Text = "BARCODE_SUPPLIER"
ElseIf Frm56.CBB5 = "Code 1" Then
    Frm56.L7_Text = "BARCODE_CODE1"
ElseIf Frm56.CBB5 = "Code 2" Then
    Frm56.L7_Text = "BARCODE_CODE2"
ElseIf Frm56.CBB5 = "Barcode" Then
    Frm56.L7_Text = "BARCODE_BARCODE"
ElseIf Frm56.CBB5 = "Berat Riyal" Then
    Frm56.L7_Text = "BARCODE_RIYAL"
ElseIf Frm56.CBB5 = "Harga" Then
    Frm56.L7_Text = "BARCODE_HARGA"
ElseIf Frm56.CBB5 = "Modal" Then
    Frm56.L7_Text = "BARCODE_MODAL"
ElseIf Frm56.CBB5 = "Diamond" Then
    Frm56.L7_Text = "BARCODE_DIAMOND"
ElseIf Frm56.CBB5 = "Design" Then
    Frm56.L7_Text = "BARCODE_DESIGN"
ElseIf Frm56.CBB5 = "No Data" Then
    Frm56.L7_Text = "No Data"
Else
    Frm56.L7_Text = vbNullString
End If

Exit Sub

logging:

LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " Frm56 : CBB5_Click" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
Call log_rekod

Resume Next
End Sub
Private Sub CBB6_Click()
'On Error GoTo logging:
If Frm56.CBB6 = "Berat" Then
    Frm56.L8_Text = "BARCODE_BERAT"
ElseIf Frm56.CBB6 = "Upah Modal" Then
    Frm56.L8_Text = "BARCODE_UPAH"
ElseIf Frm56.CBB6 = "Upah Jualan" Then
    Frm56.L8_Text = "BARCODE_UPAH2"
ElseIf Frm56.CBB6 = "Purity" Then
    Frm56.L8_Text = "BARCODE_PURITY"
ElseIf Frm56.CBB6 = "Panjang" Then
    Frm56.L8_Text = "BARCODE_Panjang"
ElseIf Frm56.CBB6 = "Lebar" Then
    Frm56.L8_Text = "BARCODE_Lebar"
ElseIf Frm56.CBB6 = "Saiz" Then
    Frm56.L8_Text = "BARCODE_Saiz"
ElseIf Frm56.CBB6 = "Dulang" Then
    Frm56.L8_Text = "BARCODE_DULANG"
ElseIf Frm56.CBB6 = "Supplier" Then
    Frm56.L8_Text = "BARCODE_SUPPLIER"
ElseIf Frm56.CBB6 = "Code 1" Then
    Frm56.L8_Text = "BARCODE_CODE1"
ElseIf Frm56.CBB6 = "Code 2" Then
    Frm56.L8_Text = "BARCODE_CODE2"
ElseIf Frm56.CBB6 = "Barcode" Then
    Frm56.L8_Text = "BARCODE_BARCODE"
ElseIf Frm56.CBB6 = "Berat Riyal" Then
    Frm56.L8_Text = "BARCODE_RIYAL"
ElseIf Frm56.CBB6 = "Harga" Then
    Frm56.L8_Text = "BARCODE_HARGA"
ElseIf Frm56.CBB6 = "Modal" Then
    Frm56.L8_Text = "BARCODE_MODAL"
ElseIf Frm56.CBB6 = "Diamond" Then
    Frm56.L8_Text = "BARCODE_DIAMOND"
ElseIf Frm56.CBB6 = "Design" Then
    Frm56.L8_Text = "BARCODE_DESIGN"
ElseIf Frm56.CBB6 = "No Data" Then
    Frm56.L8_Text = "No Data"
Else
    Frm56.L8_Text = vbNullString
End If

Exit Sub

logging:

LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " Frm56 : CBB6_Click" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
Call log_rekod

Resume Next
End Sub
Private Sub CBB7_Click()
'On Error GoTo logging:
If Frm56.CBB7 = "Berat" Then
    Frm56.L9_Text = "BARCODE_BERAT"
ElseIf Frm56.CBB7 = "Upah Modal" Then
    Frm56.L9_Text = "BARCODE_UPAH"
ElseIf Frm56.CBB7 = "Upah Jualan" Then
    Frm56.L9_Text = "BARCODE_UPAH2"
ElseIf Frm56.CBB7 = "Purity" Then
    Frm56.L9_Text = "BARCODE_PURITY"
ElseIf Frm56.CBB7 = "Panjang" Then
    Frm56.L9_Text = "BARCODE_Panjang"
ElseIf Frm56.CBB7 = "Lebar" Then
    Frm56.L9_Text = "BARCODE_Lebar"
ElseIf Frm56.CBB7 = "Saiz" Then
    Frm56.L9_Text = "BARCODE_Saiz"
ElseIf Frm56.CBB7 = "Dulang" Then
    Frm56.L9_Text = "BARCODE_DULANG"
ElseIf Frm56.CBB7 = "Supplier" Then
    Frm56.L9_Text = "BARCODE_SUPPLIER"
ElseIf Frm56.CBB7 = "Code 1" Then
    Frm56.L9_Text = "BARCODE_CODE1"
ElseIf Frm56.CBB7 = "Code 2" Then
    Frm56.L9_Text = "BARCODE_CODE2"
ElseIf Frm56.CBB7 = "Barcode" Then
    Frm56.L9_Text = "BARCODE_BARCODE"
ElseIf Frm56.CBB7 = "Berat Riyal" Then
    Frm56.L9_Text = "BARCODE_RIYAL"
ElseIf Frm56.CBB7 = "Harga" Then
    Frm56.L9_Text = "BARCODE_HARGA"
ElseIf Frm56.CBB7 = "Modal" Then
    Frm56.L9_Text = "BARCODE_MODAL"
ElseIf Frm56.CBB7 = "Diamond" Then
    Frm56.L9_Text = "BARCODE_DIAMOND"
ElseIf Frm56.CBB7 = "Design" Then
    Frm56.L9_Text = "BARCODE_DESIGN"
ElseIf Frm56.CBB7 = "No Data" Then
    Frm56.L9_Text = "No Data"
Else
    Frm56.L9_Text = vbNullString
End If

Exit Sub

logging:

LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " Frm56 : CBB7_Click" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
Call log_rekod

Resume Next
End Sub
Private Sub CBB8_Click()
'On Error GoTo logging:
If Frm56.CBB8 = "Berat" Then
    Frm56.L10_Text = "BARCODE_BERAT"
ElseIf Frm56.CBB8 = "Upah Modal" Then
    Frm56.L10_Text = "BARCODE_UPAH"
ElseIf Frm56.CBB8 = "Upah Jualan" Then
    Frm56.L10_Text = "BARCODE_UPAH2"
ElseIf Frm56.CBB8 = "Purity" Then
    Frm56.L10_Text = "BARCODE_PURITY"
ElseIf Frm56.CBB8 = "Panjang" Then
    Frm56.L10_Text = "BARCODE_Panjang"
ElseIf Frm56.CBB8 = "Lebar" Then
    Frm56.L10_Text = "BARCODE_Lebar"
ElseIf Frm56.CBB8 = "Saiz" Then
    Frm56.L10_Text = "BARCODE_Saiz"
ElseIf Frm56.CBB8 = "Dulang" Then
    Frm56.L10_Text = "BARCODE_DULANG"
ElseIf Frm56.CBB8 = "Supplier" Then
    Frm56.L10_Text = "BARCODE_SUPPLIER"
ElseIf Frm56.CBB8 = "Code 1" Then
    Frm56.L10_Text = "BARCODE_CODE1"
ElseIf Frm56.CBB8 = "Code 2" Then
    Frm56.L10_Text = "BARCODE_CODE2"
ElseIf Frm56.CBB8 = "Barcode" Then
    Frm56.L10_Text = "BARCODE_BARCODE"
ElseIf Frm56.CBB8 = "Berat Riyal" Then
    Frm56.L10_Text = "BARCODE_RIYAL"
ElseIf Frm56.CBB8 = "Harga" Then
    Frm56.L10_Text = "BARCODE_HARGA"
ElseIf Frm56.CBB8 = "Modal" Then
    Frm56.L10_Text = "BARCODE_MODAL"
ElseIf Frm56.CBB8 = "Diamond" Then
    Frm56.L10_Text = "BARCODE_DIAMOND"
ElseIf Frm56.CBB8 = "Design" Then
    Frm56.L10_Text = "BARCODE_DESIGN"
ElseIf Frm56.CBB8 = "No Data" Then
    Frm56.L10_Text = "No Data"
Else
    Frm56.L10_Text = vbNullString
End If

Exit Sub

logging:

LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " Frm56 : CBB8_Click" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
Call log_rekod

Resume Next
End Sub
Private Sub CBB9_Click()
'On Error GoTo logging:
If Frm56.CBB9 = "Berat" Then
    Frm56.L11_Text = "BARCODE_BERAT"
ElseIf Frm56.CBB9 = "Upah Modal" Then
    Frm56.L11_Text = "BARCODE_UPAH"
ElseIf Frm56.CBB9 = "Upah Jualan" Then
    Frm56.L11_Text = "BARCODE_UPAH2"
ElseIf Frm56.CBB9 = "Purity" Then
    Frm56.L11_Text = "BARCODE_PURITY"
ElseIf Frm56.CBB9 = "Panjang" Then
    Frm56.L11_Text = "BARCODE_Panjang"
ElseIf Frm56.CBB9 = "Lebar" Then
    Frm56.L11_Text = "BARCODE_Lebar"
ElseIf Frm56.CBB9 = "Saiz" Then
    Frm56.L11_Text = "BARCODE_Saiz"
ElseIf Frm56.CBB9 = "Dulang" Then
    Frm56.L11_Text = "BARCODE_DULANG"
ElseIf Frm56.CBB9 = "Supplier" Then
    Frm56.L11_Text = "BARCODE_SUPPLIER"
ElseIf Frm56.CBB9 = "Code 1" Then
    Frm56.L11_Text = "BARCODE_CODE1"
ElseIf Frm56.CBB9 = "Code 2" Then
    Frm56.L11_Text = "BARCODE_CODE2"
ElseIf Frm56.CBB9 = "Barcode" Then
    Frm56.L11_Text = "BARCODE_BARCODE"
ElseIf Frm56.CBB9 = "Berat Riyal" Then
    Frm56.L11_Text = "BARCODE_RIYAL"
ElseIf Frm56.CBB9 = "Harga" Then
    Frm56.L11_Text = "BARCODE_HARGA"
ElseIf Frm56.CBB9 = "Modal" Then
    Frm56.L11_Text = "BARCODE_MODAL"
ElseIf Frm56.CBB9 = "Diamond" Then
    Frm56.L11_Text = "BARCODE_DIAMOND"
ElseIf Frm56.CBB9 = "Design" Then
    Frm56.L11_Text = "BARCODE_DESIGN"
ElseIf Frm56.CBB9 = "No Data" Then
    Frm56.L11_Text = "No Data"
Else
    Frm56.L11_Text = vbNullString
End If

Exit Sub

logging:

LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " Frm56 : CBB9_Click" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
Call log_rekod

Resume Next
End Sub
Private Sub CBB14_Click()
'On Error GoTo logging:
If Frm56.CBB14 = "Berat" Then
    Frm56.L12_Text = "BARCODE_BERAT"
ElseIf Frm56.CBB14 = "Upah Modal" Then
    Frm56.L12_Text = "BARCODE_UPAH"
ElseIf Frm56.CBB14 = "Upah Jualan" Then
    Frm56.L12_Text = "BARCODE_UPAH2"
ElseIf Frm56.CBB14 = "Purity" Then
    Frm56.L12_Text = "BARCODE_PURITY"
ElseIf Frm56.CBB14 = "Panjang" Then
    Frm56.L12_Text = "BARCODE_Panjang"
ElseIf Frm56.CBB14 = "Lebar" Then
    Frm56.L12_Text = "BARCODE_Lebar"
ElseIf Frm56.CBB14 = "Saiz" Then
    Frm56.L12_Text = "BARCODE_Saiz"
ElseIf Frm56.CBB14 = "Dulang" Then
    Frm56.L12_Text = "BARCODE_DULANG"
ElseIf Frm56.CBB14 = "Supplier" Then
    Frm56.L12_Text = "BARCODE_SUPPLIER"
ElseIf Frm56.CBB14 = "Code 1" Then
    Frm56.L12_Text = "BARCODE_CODE1"
ElseIf Frm56.CBB14 = "Code 2" Then
    Frm56.L12_Text = "BARCODE_CODE2"
ElseIf Frm56.CBB14 = "Barcode" Then
    Frm56.L12_Text = "BARCODE_BARCODE"
ElseIf Frm56.CBB14 = "Berat Riyal" Then
    Frm56.L12_Text = "BARCODE_RIYAL"
ElseIf Frm56.CBB14 = "Harga" Then
    Frm56.L12_Text = "BARCODE_HARGA"
ElseIf Frm56.CBB14 = "Modal" Then
    Frm56.L12_Text = "BARCODE_MODAL"
ElseIf Frm56.CBB14 = "Diamond" Then
    Frm56.L12_Text = "BARCODE_DIAMOND"
ElseIf Frm56.CBB14 = "Design" Then
    Frm56.L12_Text = "BARCODE_DESIGN"
ElseIf Frm56.CBB14 = "No Data" Then
    Frm56.L12_Text = "No Data"
Else
    Frm56.L12_Text = vbNullString
End If

Exit Sub

logging:

LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " Frm56 : CBB14_Click" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
Call log_rekod

Resume Next
End Sub
Private Sub CBB15_Click()
'On Error GoTo logging:
If Frm56.CBB15 = "Berat" Then
    Frm56.L13_Text = "BARCODE_BERAT"
ElseIf Frm56.CBB15 = "Upah Modal" Then
    Frm56.L13_Text = "BARCODE_UPAH"
ElseIf Frm56.CBB15 = "Upah Jualan" Then
    Frm56.L13_Text = "BARCODE_UPAH2"
ElseIf Frm56.CBB15 = "Purity" Then
    Frm56.L13_Text = "BARCODE_PURITY"
ElseIf Frm56.CBB15 = "Panjang" Then
    Frm56.L13_Text = "BARCODE_Panjang"
ElseIf Frm56.CBB15 = "Lebar" Then
    Frm56.L13_Text = "BARCODE_Lebar"
ElseIf Frm56.CBB15 = "Saiz" Then
    Frm56.L13_Text = "BARCODE_Saiz"
ElseIf Frm56.CBB15 = "Dulang" Then
    Frm56.L13_Text = "BARCODE_DULANG"
ElseIf Frm56.CBB15 = "Supplier" Then
    Frm56.L13_Text = "BARCODE_SUPPLIER"
ElseIf Frm56.CBB15 = "Code 1" Then
    Frm56.L13_Text = "BARCODE_CODE1"
ElseIf Frm56.CBB15 = "Code 2" Then
    Frm56.L13_Text = "BARCODE_CODE2"
ElseIf Frm56.CBB15 = "Barcode" Then
    Frm56.L13_Text = "BARCODE_BARCODE"
ElseIf Frm56.CBB15 = "Berat Riyal" Then
    Frm56.L13_Text = "BARCODE_RIYAL"
ElseIf Frm56.CBB15 = "Harga" Then
    Frm56.L13_Text = "BARCODE_HARGA"
ElseIf Frm56.CBB15 = "Modal" Then
    Frm56.L13_Text = "BARCODE_MODAL"
ElseIf Frm56.CBB15 = "Diamond" Then
    Frm56.L13_Text = "BARCODE_DIAMOND"
ElseIf Frm56.CBB15 = "Design" Then
    Frm56.L13_Text = "BARCODE_DESIGN"
ElseIf Frm56.CBB15 = "No Data" Then
    Frm56.L13_Text = "No Data"
Else
    Frm56.L13_Text = vbNullString
End If

Exit Sub

logging:

LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " Frm56 : CBB15_Click" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
Call log_rekod

Resume Next
End Sub
Private Sub CBB16_Click()
'On Error GoTo logging:
If Frm56.CBB16 = "Berat" Then
    Frm56.L14_Text = "BARCODE_BERAT"
ElseIf Frm56.CBB16 = "Upah Modal" Then
    Frm56.L14_Text = "BARCODE_UPAH"
ElseIf Frm56.CBB16 = "Upah Jualan" Then
    Frm56.L14_Text = "BARCODE_UPAH2"
ElseIf Frm56.CBB16 = "Purity" Then
    Frm56.L14_Text = "BARCODE_PURITY"
ElseIf Frm56.CBB16 = "Panjang" Then
    Frm56.L14_Text = "BARCODE_Panjang"
ElseIf Frm56.CBB16 = "Lebar" Then
    Frm56.L14_Text = "BARCODE_Lebar"
ElseIf Frm56.CBB16 = "Saiz" Then
    Frm56.L14_Text = "BARCODE_Saiz"
ElseIf Frm56.CBB16 = "Dulang" Then
    Frm56.L14_Text = "BARCODE_DULANG"
ElseIf Frm56.CBB16 = "Supplier" Then
    Frm56.L14_Text = "BARCODE_SUPPLIER"
ElseIf Frm56.CBB16 = "Code 1" Then
    Frm56.L14_Text = "BARCODE_CODE1"
ElseIf Frm56.CBB16 = "Code 2" Then
    Frm56.L14_Text = "BARCODE_CODE2"
ElseIf Frm56.CBB16 = "Barcode" Then
    Frm56.L14_Text = "BARCODE_BARCODE"
ElseIf Frm56.CBB16 = "Berat Riyal" Then
    Frm56.L14_Text = "BARCODE_RIYAL"
ElseIf Frm56.CBB16 = "Harga" Then
    Frm56.L14_Text = "BARCODE_HARGA"
ElseIf Frm56.CBB16 = "Modal" Then
    Frm56.L14_Text = "BARCODE_MODAL"
ElseIf Frm56.CBB16 = "Diamond" Then
    Frm56.L14_Text = "BARCODE_DIAMOND"
ElseIf Frm56.CBB16 = "Design" Then
    Frm56.L14_Text = "BARCODE_DESIGN"
ElseIf Frm56.CBB16 = "No Data" Then
    Frm56.L14_Text = "No Data"
Else
    Frm56.L14_Text = vbNullString
End If

Exit Sub

logging:

LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " Frm56 : CBB16_Click" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
Call log_rekod

Resume Next
End Sub
Private Sub CMD1_Click()
'On Error GoTo logging:
Dim Data_Err(5)
DATA_SAVE = 0

GoTo skip_this:
If Frm56.CBB10 = vbNullString Then
    x = x + 1
    Data_Err(x) = "Sila pilih [Saiz tulisan barisan pertama]"
End If
If Frm56.CBB11 = vbNullString Then
    x = x + 1
    Data_Err(x) = "Sila pilih [Saiz tulisan barisan kedua]"
End If
If Frm56.CBB12 = vbNullString Then
    x = x + 1
    Data_Err(x) = "Sila pilih [Saiz tulisan barisan ketiga]"
End If
If Frm56.CBB13 = vbNullString Then
    x = x + 1
    Data_Err(x) = "Sila pilih [Saiz tulisan barisan keempat]"
End If
If Frm56.CB14 = 0 And Frm56.CB15 = 0 Then
    x = x + 1
    Data_Err(x) = "Sila pilih jenis barcode label."
End If
If Frm56.CBB20 = vbNullString Then
    x = x + 1
    Data_Err(x) = "Sila pilih [Font Type tulisan barisan pertama]"
End If
If Frm56.CBB21 = vbNullString Then
    x = x + 1
    Data_Err(x) = "Sila pilih [Font Type tulisan barisan kedua]"
End If
If Frm56.CBB22 = vbNullString Then
    x = x + 1
    Data_Err(x) = "Sila pilih [Font Type tulisan barisan ketiga]"
End If
If Frm56.CBB23 = vbNullString Then
    x = x + 1
    Data_Err(x) = "Sila pilih [Font Type tulisan barisan keempat]"
End If
skip_this:
If x <> 0 Then
    Frm6.Show
    Frm6.Pic1.Cls
    For Y = 1 To x
        Frm6.Pic1.Print Y & " - " & Data_Err(Y)
    Next Y
    Exit Sub
Else

    Note = "Simpan tetapan ini ?" & vbCrLf & _
            "Sistem mungkin mengambil sedikit masa untuk menyimpan tetapan ini." & vbCrLf & _
            vbNullString & vbCrLf & _
            "Teruskan?"
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbYes Then

        If MDI_frm1.L20_Text = "Semua cawangan" Then
            LM_KEDAI = "HQ"
        Else
            LM_KEDAI = MDI_frm1.L20_Text
        End If
        LM_NOW = Now
        
        '#########Layout Barcode###########
LM_CONN = 1
re_conn_1:
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from layout_barcode where perkara='" & G_JENIS_BARCODE_PRINTER & "'", cn, adOpenKeyset, adLockOptimistic

        If Not rs.EOF Then
            If Frm56.L15_Text = "0" Then '0 : Type A , 1 : Type B
                rs!a_line = Frm56.CB1.Value & "," & Frm56.CB2.Value & "," & Frm56.CB3.Value & "," & Frm56.CB4.Value
                rs!a_data = Frm56.L3_Text & "," & Frm56.L4_Text & "," & Frm56.L12_Text & "," & Frm56.L5_Text & "," & Frm56.L6_Text & "," & Frm56.L7_Text & "," & Frm56.L8_Text & "," & Frm56.L9_Text & "," & Frm56.L13_Text & "," & Frm56.L10_Text & "," & Frm56.L11_Text & "," & Frm56.L14_Text & ","
                rs!a_pre_data = Frm56.CBB1 & "," & Frm56.CBB2 & "," & Frm56.CBB14 & "," & Frm56.CBB3 & "," & Frm56.CBB4 & "," & Frm56.CBB5 & "," & Frm56.CBB6 & "," & Frm56.CBB7 & "," & Frm56.CBB15 & "," & Frm56.CBB8 & "," & Frm56.CBB9 & "," & Frm56.CBB16 & ","
                LM_JENIS = "Type A"
            ElseIf Frm56.L15_Text = "1" Then '0 : Type A , 1 : Type B
                rs!b_line = Frm56.CB1.Value & "," & Frm56.CB2.Value & "," & Frm56.CB3.Value & "," & Frm56.CB4.Value
                rs!b_data = Frm56.L3_Text & "," & Frm56.L4_Text & "," & Frm56.L12_Text & "," & Frm56.L5_Text & "," & Frm56.L6_Text & "," & Frm56.L7_Text & "," & Frm56.L8_Text & "," & Frm56.L9_Text & "," & Frm56.L13_Text & "," & Frm56.L10_Text & "," & Frm56.L11_Text & "," & Frm56.L14_Text & ","
                rs!b_pre_data = Frm56.CBB1 & "," & Frm56.CBB2 & "," & Frm56.CBB14 & "," & Frm56.CBB3 & "," & Frm56.CBB4 & "," & Frm56.CBB5 & "," & Frm56.CBB6 & "," & Frm56.CBB7 & "," & Frm56.CBB15 & "," & Frm56.CBB8 & "," & Frm56.CBB9 & "," & Frm56.CBB16 & ","
                LM_JENIS = "Type B"
            End If
            DATA_SAVE = 1
            rs.Update
        End If
        
        rs.Close
        Set rs = Nothing
        
        If DATA_SAVE = 1 Then
            LogAct_Memory = "Tetapan layout barcode. (" & LM_JENIS & ")."
            LogDate_Memory = LM_NOW
            Call UpdateLog_Database
            
            If LM_KEDAI = MDI_frm1.L20_Text Then Call setting_barcode
            MsgBox "Tetapan telah berjaya disimpan.", vbInformation, "Info"
        End If
    End If
End If

Exit Sub

logging:

LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " Frm56 : CMD1_Click" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
Call log_rekod

If LM_ERR_NO = "3704" Or LM_ERR_NO = "-2147467259" Or LM_ERR_NO = "-2147217887" Then
    Call Main
    
    If LM_CONN = 1 Then
        Resume re_conn_1:
    End If
Else
    Resume Next
End If
End Sub
Private Sub CMD2_Click()
'On Error GoTo logging:
Dim Data_Err(5)
DATA_SAVE = 0

If x <> 0 Then
    Frm6.Show
    Frm6.Pic1.Cls
    For Y = 1 To x
        Frm6.Pic1.Print Y & " - " & Data_Err(Y)
    Next Y
    Exit Sub
Else

    Note = "Simpan tetapan ini ?" & vbCrLf & _
            "Sistem mungkin mengambil sedikit masa untuk menyimpan tetapan ini." & vbCrLf & _
            vbNullString & vbCrLf & _
            "Teruskan?"
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbYes Then
        LM_FONT_SIZE = vbNullString
        LM_FONT_TYPE = vbNullString
        LM_POS_X = vbNullString
        LM_POS_Y = vbNullString
        LM_BOLD = vbNullString
        LM_ITALIC = vbNullString
            
        For Y = 0 To 11
            LM_FONT_SIZE = LM_FONT_SIZE & Frm56.TB1(Y) & ","
            LM_FONT_TYPE = LM_FONT_TYPE & Frm56.TB2(Y) & ","
            LM_POS_X = LM_POS_X & Frm56.TB3(Y) & ","
            LM_POS_Y = LM_POS_Y & Frm56.TB4(Y) & ","
            LM_BOLD = LM_BOLD & Frm56.CB10(Y).Value & ","
            LM_ITALIC = LM_ITALIC & Frm56.CB11(Y).Value & ","
        Next Y
        LM_NOW = Now
        '#########Layout Barcode###########
LM_CONN = 1
re_conn_1:
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from layout_barcode where perkara='" & G_ID & "'", cn, adOpenKeyset, adLockOptimistic
    
        If Not rs.EOF Then
            If Frm56.L15_Text = "0" Then '0 : Type A , 1 : Type B , 2 : Type C
                rs!a_font_size = LM_FONT_SIZE
                rs!a_font_type = LM_FONT_TYPE
                rs!a_position_x = LM_POS_X
                rs!a_position_y = LM_POS_Y
                rs!a_bold = LM_BOLD
                rs!a_italic = LM_ITALIC
                LM_JENIS = "Type A"
            ElseIf Frm56.L15_Text = "1" Then '0 : Type A , 1 : Type B , 2 : Type C
                rs!b_font_size = LM_FONT_SIZE
                rs!b_font_type = LM_FONT_TYPE
                rs!b_position_x = LM_POS_X
                rs!b_position_y = LM_POS_Y
                rs!b_bold = LM_BOLD
                rs!b_italic = LM_ITALIC
                LM_JENIS = "Type B"
            ElseIf Frm56.L15_Text = "2" Then '0 : Type A , 1 : Type B , 2 : Type C
                rs!c_font_size = LM_FONT_SIZE
                rs!c_font_type = LM_FONT_TYPE
                rs!c_position_x = LM_POS_X
                rs!c_position_y = LM_POS_Y
                rs!c_bold = LM_BOLD
                rs!c_italic = LM_ITALIC
                LM_JENIS = "Type C"
            End If
            rs.Update
            DATA_SAVE = 1
        End If
        
        rs.Close
        Set rs = Nothing
        
        If DATA_SAVE = 1 Then
            LogAct_Memory = "Tetapan element SKU/Label. Model [" & G_ID & "] , Type [" & LM_JENIS & "]."
            LogDate_Memory = LM_NOW
            Call UpdateLog_Database
            'If LM_KEDAI = MDI_frm1.L20_Text Then Call setting_barcode
            Call setting_barcode
            
            MsgBox "Tetapan telah berjaya disimpan.", vbInformation, "Info"
        End If
    End If
End If

Exit Sub

logging:

LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " Frm56 : CMD2_Click" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
Call log_rekod

If LM_ERR_NO = "3704" Or LM_ERR_NO = "-2147467259" Or LM_ERR_NO = "-2147217887" Then
    Call Main
    If LM_CONN = 1 Then
        Resume re_conn_1:
    End If
Else
    Resume Next
End If
End Sub

Private Sub CMD3_Click()
'On Error GoTo logging:
Dim Data_Err(5)
DATA_SAVE = 0

If Frm56.CB14 = 0 And Frm56.CB15 = 0 And Frm56.CB16 = 0 Then
    x = x + 1
    Data_Err(x) = "Sila pilih jenis barcode label."
End If

If x <> 0 Then
    Frm6.Show
    Frm6.Pic1.Cls
    For Y = 1 To x
        Frm6.Pic1.Print Y & " - " & Data_Err(Y)
    Next Y
    Exit Sub
Else

    Note = "Simpan tetapan ini ?" & vbCrLf & _
            "Sistem mungkin mengambil sedikit masa untuk menyimpan tetapan ini." & vbCrLf & _
            vbNullString & vbCrLf & _
            "Teruskan?"
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbYes Then
        
        If Frm56.CB14 = 1 Then
            LM_TYPE_BARCODE = 0
            LM_JENIS = "Type A"
        ElseIf Frm56.CB15 = 1 Then
            LM_TYPE_BARCODE = 1
            LM_JENIS = "Type B"
        ElseIf Frm56.CB16 = 1 Then
            LM_TYPE_BARCODE = 2
            LM_JENIS = "Type C"
        End If
        
        LM_NOW = Now
        
        '#########Layout Barcode###########
LM_CONN = 1
re_conn_1:
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        strsql = "UPDATE layout_barcode set barcode_type='" & LM_TYPE_BARCODE & "'"
        
        Set rs = cn.Execute(strsql)
        Set rs = Nothing
        
        DATA_SAVE = 1
        
        If DATA_SAVE = 1 Then
            LogAct_Memory = "Tetapan jenis barcode -> " & LM_JENIS & "."
            LogDate_Memory = LM_NOW
            Call UpdateLog_Database
            Call setting_barcode
            
            MsgBox "Tetapan telah berjaya disimpan.", vbInformation, "Info"
        End If
    End If
End If

Exit Sub

logging:

LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " Frm56 : CMD3_Click" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
Call log_rekod

If LM_ERR_NO = "3704" Or LM_ERR_NO = "-2147467259" Or LM_ERR_NO = "-2147217887" Then
    Call Main
    If LM_CONN = 1 Then
        Resume re_conn_1:
    End If
Else
    Resume Next
End If
End Sub

Private Sub Form_Load()
'On Error GoTo logging:
Frm56.Picture = MDI_frm1.Picture

Frm56.LV1.ListItems.Clear

With Frm56.LV1
    Set .SmallIcons = Frm56.ImageList4
    Set .Icons = Frm56.ImageList4

    .ListItems.Add , "B.Kemas", "B.Kemas", 43
    .ListItems.Add , "B.Permata", "B.Permata", 44
    If MDI_frm1.L4_Text = "Developer" Then
        .ListItems.Add , "Type A", "Type A", 34
        .ListItems.Add , "Type B", "Type B", 34
        .ListItems.Add , "Type C", "Type C", 34
        .ListItems.Add , "Developer", "Developer", 40
    End If
End With

Call frm56_font_type

Exit Sub

logging:

LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " Frm56 : Form_Load" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
Call log_rekod

Resume Next
End Sub

Private Sub LV1_Click()
'On Error Resume Next
LM_KEY = Frm56.LV1.SelectedItem.Key

If LM_KEY = "B.Kemas" Then
    Call frm56_initial_frame
    Frm56.L15_Text = "0" '0 : Type A , 1 : Type B
    Call Frm56_ListItem_BK
    Frm56.Frame2.Caption = "Barang Kemas"
    Frm56.Frame2.Visible = True
ElseIf LM_KEY = "B.Permata" Then
    Call frm56_initial_frame
    Frm56.L15_Text = "1" '0 : Type A , 1 : Type B
    Call Frm56_ListItem_permata
    Frm56.Frame2.Caption = "Barang Permata"
    Frm56.Frame2.Visible = True
ElseIf LM_KEY = "Type A" Or LM_KEY = "Type B" Or LM_KEY = "Type C" Then
    Call frm56_initial_frame
    Call frm56_initial_frame2
    If LM_KEY = "Type A" Then
        Frm56.Frame3.Caption = "Type A"
        Frm56.L15_Text = "0" '0 : Type A , 1 : Type B , 2 : Type C
    ElseIf LM_KEY = "Type B" Then
        Frm56.L15_Text = "1" '0 : Type A , 1 : Type B , 2 : Type C
        Frm56.Frame3.Caption = "Type B"
    ElseIf LM_KEY = "Type C" Then
        Frm56.L15_Text = "2" '0 : Type A , 1 : Type B , 2 : Type C
        Frm56.Frame3.Caption = "Type C"
    End If
    Call frm56_senarai_printer
    Frm56.Frame3.Visible = True
ElseIf LM_KEY = "Type CA" Then
    Call frm56_initial_frame
    Frm56.L15_Text = "0" '0 : Type A , 1 : Type B
    Call frm56_reset_element
    Frm56.Frame3.Caption = "Type A"
    Frm56.Frame3.Visible = True
ElseIf LM_KEY = "Type CB" Then
    Call frm56_initial_frame
    Frm56.L15_Text = "1" '0 : Type A , 1 : Type B
    Call frm56_reset_element
    Frm56.Frame3.Caption = "Type B"
    Frm56.Frame3.Visible = True
ElseIf LM_KEY = "Developer" Then
    Call frm56_initial_frame
    If G_L_JENIS_BARCODE = 0 Then
        Frm56.CB14 = 1
        Frm56.CB15 = 0
        Frm56.CB16 = 0
    ElseIf G_L_JENIS_BARCODE = 1 Then
        Frm56.CB14 = 0
        Frm56.CB15 = 1
        Frm56.CB16 = 0
    ElseIf G_L_JENIS_BARCODE = 2 Then
        Frm56.CB14 = 0
        Frm56.CB15 = 0
        Frm56.CB16 = 1
    End If
    Frm56.Frame1.Visible = True
End If
End Sub
Private Sub LV2_Click()
'On Error Resume Next
LM_KEY = Frm56.LV2.SelectedItem.Key

If LM_KEY <> vbNullString Then
    Call frm56_initial_frame2
    G_ID = LM_KEY
    Frm56.Frame5.Caption = LM_KEY
    Call frm56_recall_setting_barcode
    Frm56.Frame5.Visible = True
End If
End Sub
Private Sub Tmr1_Timer()
'On Error Resume Next
Frm56.L1_Text = DateTime.Date
Frm56.L2_Text = DateTime.Time$
End Sub
Private Sub Frm56_ListItem_permata()
'On Error GoTo logging:
Frm56.CBB1.Clear
Frm56.CBB2.Clear
Frm56.CBB3.Clear
Frm56.CBB4.Clear
Frm56.CBB5.Clear
Frm56.CBB6.Clear
Frm56.CBB7.Clear
Frm56.CBB8.Clear
Frm56.CBB9.Clear
Frm56.CBB14.Clear
Frm56.CBB15.Clear
Frm56.CBB16.Clear

With Frm56.CBB1
    .AddItem "No Data"
    .AddItem "Barcode"
    .AddItem "Code 1"
    .AddItem "Code 2"
    .AddItem "Design"
    .AddItem "Diamond"
    .AddItem "Dulang"
    .AddItem "Harga"
    .AddItem "Lebar"
    .AddItem "Modal"
    .AddItem "Purity"
    .AddItem "Panjang"
    .AddItem "Saiz"
    .AddItem "Supplier"
End With

With Frm56.CBB2
    .AddItem "No Data"
    .AddItem "Barcode"
    .AddItem "Code 1"
    .AddItem "Code 2"
    .AddItem "Design"
    .AddItem "Diamond"
    .AddItem "Dulang"
    .AddItem "Harga"
    .AddItem "Lebar"
    .AddItem "Modal"
    .AddItem "Purity"
    .AddItem "Panjang"
    .AddItem "Saiz"
    .AddItem "Supplier"
End With

With Frm56.CBB3
    .AddItem "No Data"
    .AddItem "Barcode"
    .AddItem "Code 1"
    .AddItem "Code 2"
    .AddItem "Design"
    .AddItem "Diamond"
    .AddItem "Dulang"
    .AddItem "Harga"
    .AddItem "Lebar"
    .AddItem "Modal"
    .AddItem "Purity"
    .AddItem "Panjang"
    .AddItem "Saiz"
    .AddItem "Supplier"
End With

With Frm56.CBB4
    .AddItem "No Data"
    .AddItem "Barcode"
    .AddItem "Code 1"
    .AddItem "Code 2"
    .AddItem "Design"
    .AddItem "Diamond"
    .AddItem "Dulang"
    .AddItem "Harga"
    .AddItem "Lebar"
    .AddItem "Modal"
    .AddItem "Purity"
    .AddItem "Panjang"
    .AddItem "Saiz"
    .AddItem "Supplier"
End With

With Frm56.CBB5
    .AddItem "No Data"
    .AddItem "Barcode"
    .AddItem "Code 1"
    .AddItem "Code 2"
    .AddItem "Design"
    .AddItem "Diamond"
    .AddItem "Dulang"
    .AddItem "Harga"
    .AddItem "Lebar"
    .AddItem "Modal"
    .AddItem "Purity"
    .AddItem "Panjang"
    .AddItem "Saiz"
    .AddItem "Supplier"
End With

With Frm56.CBB6
    .AddItem "No Data"
    .AddItem "Barcode"
    .AddItem "Code 1"
    .AddItem "Code 2"
    .AddItem "Design"
    .AddItem "Diamond"
    .AddItem "Dulang"
    .AddItem "Harga"
    .AddItem "Lebar"
    .AddItem "Modal"
    .AddItem "Purity"
    .AddItem "Panjang"
    .AddItem "Saiz"
    .AddItem "Supplier"
End With

With Frm56.CBB7
    .AddItem "No Data"
    .AddItem "Barcode"
    .AddItem "Code 1"
    .AddItem "Code 2"
    .AddItem "Design"
    .AddItem "Diamond"
    .AddItem "Dulang"
    .AddItem "Harga"
    .AddItem "Lebar"
    .AddItem "Modal"
    .AddItem "Purity"
    .AddItem "Panjang"
    .AddItem "Saiz"
    .AddItem "Supplier"
End With

With Frm56.CBB8
    .AddItem "No Data"
    .AddItem "Barcode"
    .AddItem "Code 1"
    .AddItem "Code 2"
    .AddItem "Design"
    .AddItem "Diamond"
    .AddItem "Dulang"
    .AddItem "Harga"
    .AddItem "Lebar"
    .AddItem "Modal"
    .AddItem "Purity"
    .AddItem "Panjang"
    .AddItem "Saiz"
    .AddItem "Supplier"
End With

With Frm56.CBB9
    .AddItem "No Data"
    .AddItem "Barcode"
    .AddItem "Code 1"
    .AddItem "Code 2"
    .AddItem "Design"
    .AddItem "Diamond"
    .AddItem "Dulang"
    .AddItem "Harga"
    .AddItem "Lebar"
    .AddItem "Modal"
    .AddItem "Purity"
    .AddItem "Panjang"
    .AddItem "Saiz"
    .AddItem "Supplier"
End With

With Frm56.CBB14
    .AddItem "No Data"
    .AddItem "Barcode"
    .AddItem "Code 1"
    .AddItem "Code 2"
    .AddItem "Design"
    .AddItem "Diamond"
    .AddItem "Dulang"
    .AddItem "Harga"
    .AddItem "Lebar"
    .AddItem "Modal"
    .AddItem "Purity"
    .AddItem "Panjang"
    .AddItem "Saiz"
    .AddItem "Supplier"
End With

With Frm56.CBB15
    .AddItem "No Data"
    .AddItem "Barcode"
    .AddItem "Code 1"
    .AddItem "Code 2"
    .AddItem "Design"
    .AddItem "Diamond"
    .AddItem "Dulang"
    .AddItem "Harga"
    .AddItem "Lebar"
    .AddItem "Modal"
    .AddItem "Purity"
    .AddItem "Panjang"
    .AddItem "Saiz"
    .AddItem "Supplier"
End With

With Frm56.CBB16
    .AddItem "No Data"
    .AddItem "Barcode"
    .AddItem "Code 1"
    .AddItem "Code 2"
    .AddItem "Design"
    .AddItem "Diamond"
    .AddItem "Dulang"
    .AddItem "Harga"
    .AddItem "Lebar"
    .AddItem "Modal"
    .AddItem "Purity"
    .AddItem "Panjang"
    .AddItem "Saiz"
    .AddItem "Supplier"
End With

Dim Frm56_LM_BARCODE(13)
Dim Frm56_LM_Jenis(13)

Frm56.CBB10.Clear
Frm56.CBB11.Clear
Frm56.CBB12.Clear
Frm56.CBB13.Clear

Frm56.CB1.Value = Split(G_SKU_LINE(1), ",")(0)
Frm56.CB2.Value = Split(G_SKU_LINE(1), ",")(1)
Frm56.CB3.Value = Split(G_SKU_LINE(1), ",")(2)
Frm56.CB4.Value = Split(G_SKU_LINE(1), ",")(3)
 

If Split(G_SKU_PRE_DATA(1), ",")(0) <> "" Then Frm56.CBB1 = Split(G_SKU_PRE_DATA(1), ",")(0)
If Split(G_SKU_PRE_DATA(1), ",")(1) <> "" Then Frm56.CBB2 = Split(G_SKU_PRE_DATA(1), ",")(1)
If Split(G_SKU_PRE_DATA(1), ",")(2) <> "" Then Frm56.CBB14 = Split(G_SKU_PRE_DATA(1), ",")(2)
If Split(G_SKU_PRE_DATA(1), ",")(3) <> "" Then Frm56.CBB3 = Split(G_SKU_PRE_DATA(1), ",")(3)
If Split(G_SKU_PRE_DATA(1), ",")(4) <> "" Then Frm56.CBB4 = Split(G_SKU_PRE_DATA(1), ",")(4)
If Split(G_SKU_PRE_DATA(1), ",")(5) <> "" Then Frm56.CBB5 = Split(G_SKU_PRE_DATA(1), ",")(5)
If Split(G_SKU_PRE_DATA(1), ",")(6) <> "" Then Frm56.CBB6 = Split(G_SKU_PRE_DATA(1), ",")(6)
If Split(G_SKU_PRE_DATA(1), ",")(7) <> "" Then Frm56.CBB7 = Split(G_SKU_PRE_DATA(1), ",")(7)
If Split(G_SKU_PRE_DATA(1), ",")(8) <> "" Then Frm56.CBB15 = Split(G_SKU_PRE_DATA(1), ",")(8)
If Split(G_SKU_PRE_DATA(1), ",")(9) <> "" Then Frm56.CBB8 = Split(G_SKU_PRE_DATA(1), ",")(9)
If Split(G_SKU_PRE_DATA(1), ",")(10) <> "" Then Frm56.CBB9 = Split(G_SKU_PRE_DATA(1), ",")(10)
If Split(G_SKU_PRE_DATA(1), ",")(11) <> "" Then Frm56.CBB16 = Split(G_SKU_PRE_DATA(1), ",")(11)

If Split(G_SKU_DATA(1), ",")(0) <> "No Data" Then Frm56.L3_Text = Split(G_SKU_DATA(1), ",")(0)
If Split(G_SKU_DATA(1), ",")(1) <> "No Data" Then Frm56.L4_Text = Split(G_SKU_DATA(1), ",")(1)
If Split(G_SKU_DATA(1), ",")(2) <> "No Data" Then Frm56.L12_Text = Split(G_SKU_DATA(1), ",")(2)
If Split(G_SKU_DATA(1), ",")(3) <> "No Data" Then Frm56.L5_Text = Split(G_SKU_DATA(1), ",")(3)
If Split(G_SKU_DATA(1), ",")(4) <> "No Data" Then Frm56.L6_Text = Split(G_SKU_DATA(1), ",")(4)
If Split(G_SKU_DATA(1), ",")(5) <> "No Data" Then Frm56.L7_Text = Split(G_SKU_DATA(1), ",")(5)
If Split(G_SKU_DATA(1), ",")(6) <> "No Data" Then Frm56.L8_Text = Split(G_SKU_DATA(1), ",")(6)
If Split(G_SKU_DATA(1), ",")(7) <> "No Data" Then Frm56.L9_Text = Split(G_SKU_DATA(1), ",")(7)
If Split(G_SKU_DATA(1), ",")(8) <> "No Data" Then Frm56.L13_Text = Split(G_SKU_DATA(1), ",")(8)
If Split(G_SKU_DATA(1), ",")(9) <> "No Data" Then Frm56.L10_Text = Split(G_SKU_DATA(1), ",")(9)
If Split(G_SKU_DATA(1), ",")(10) <> "No Data" Then Frm56.L11_Text = Split(G_SKU_DATA(1), ",")(10)
If Split(G_SKU_DATA(1), ",")(11) <> "No Data" Then Frm56.L14_Text = Split(G_SKU_DATA(1), ",")(11)

Exit Sub

logging:

LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " Frm56 : Frm56_ListItem_permata" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
Call log_rekod

Resume Next
End Sub

Private Sub Frm56_ListItem_BK()
'On Error GoTo logging:
Frm56.CBB1.Clear
Frm56.CBB2.Clear
Frm56.CBB3.Clear
Frm56.CBB4.Clear
Frm56.CBB5.Clear
Frm56.CBB6.Clear
Frm56.CBB7.Clear
Frm56.CBB8.Clear
Frm56.CBB9.Clear
Frm56.CBB14.Clear
Frm56.CBB15.Clear
Frm56.CBB16.Clear

With Frm56.CBB1
    .AddItem "No Data"
    .AddItem "Barcode"
    .AddItem "Berat"
    .AddItem "Berat Riyal"
    .AddItem "Code 1"
    .AddItem "Code 2"
    .AddItem "Dulang"
    .AddItem "Design"
    .AddItem "Diamond"
    .AddItem "Lebar"
    .AddItem "Panjang"
    .AddItem "Purity"
    .AddItem "Upah Modal"
    .AddItem "Upah Jualan"
    .AddItem "Saiz"
    .AddItem "Supplier"
End With
    
With Frm56.CBB2
    .AddItem "No Data"
    .AddItem "Barcode"
    .AddItem "Berat"
    .AddItem "Berat Riyal"
    .AddItem "Code 1"
    .AddItem "Code 2"
    .AddItem "Dulang"
    .AddItem "Design"
    .AddItem "Diamond"
    .AddItem "Lebar"
    .AddItem "Panjang"
    .AddItem "Purity"
    .AddItem "Upah Modal"
    .AddItem "Upah Jualan"
    .AddItem "Saiz"
    .AddItem "Supplier"
End With

With Frm56.CBB3
    .AddItem "No Data"
    .AddItem "Barcode"
    .AddItem "Berat"
    .AddItem "Berat Riyal"
    .AddItem "Code 1"
    .AddItem "Code 2"
    .AddItem "Dulang"
    .AddItem "Design"
    .AddItem "Diamond"
    .AddItem "Lebar"
    .AddItem "Panjang"
    .AddItem "Purity"
    .AddItem "Upah Modal"
    .AddItem "Upah Jualan"
    .AddItem "Saiz"
    .AddItem "Supplier"
End With

With Frm56.CBB4
    .AddItem "No Data"
    .AddItem "Barcode"
    .AddItem "Berat"
    .AddItem "Berat Riyal"
    .AddItem "Code 1"
    .AddItem "Code 2"
    .AddItem "Dulang"
    .AddItem "Design"
    .AddItem "Diamond"
    .AddItem "Lebar"
    .AddItem "Panjang"
    .AddItem "Purity"
    .AddItem "Upah Modal"
    .AddItem "Upah Jualan"
    .AddItem "Saiz"
    .AddItem "Supplier"
End With

With Frm56.CBB5
    .AddItem "No Data"
    .AddItem "Barcode"
    .AddItem "Berat"
    .AddItem "Berat Riyal"
    .AddItem "Code 1"
    .AddItem "Code 2"
    .AddItem "Dulang"
    .AddItem "Design"
    .AddItem "Diamond"
    .AddItem "Lebar"
    .AddItem "Panjang"
    .AddItem "Purity"
    .AddItem "Upah Modal"
    .AddItem "Upah Jualan"
    .AddItem "Saiz"
    .AddItem "Supplier"
End With

With Frm56.CBB6
    .AddItem "No Data"
    .AddItem "Barcode"
    .AddItem "Berat"
    .AddItem "Berat Riyal"
    .AddItem "Code 1"
    .AddItem "Code 2"
    .AddItem "Dulang"
    .AddItem "Design"
    .AddItem "Diamond"
    .AddItem "Lebar"
    .AddItem "Panjang"
    .AddItem "Purity"
    .AddItem "Upah Modal"
    .AddItem "Upah Jualan"
    .AddItem "Saiz"
    .AddItem "Supplier"
End With

With Frm56.CBB7
    .AddItem "No Data"
    .AddItem "Barcode"
    .AddItem "Berat"
    .AddItem "Berat Riyal"
    .AddItem "Code 1"
    .AddItem "Code 2"
    .AddItem "Dulang"
    .AddItem "Design"
    .AddItem "Diamond"
    .AddItem "Lebar"
    .AddItem "Panjang"
    .AddItem "Purity"
    .AddItem "Upah Modal"
    .AddItem "Upah Jualan"
    .AddItem "Saiz"
    .AddItem "Supplier"
End With

With Frm56.CBB8
    .AddItem "No Data"
    .AddItem "Barcode"
    .AddItem "Berat"
    .AddItem "Berat Riyal"
    .AddItem "Code 1"
    .AddItem "Code 2"
    .AddItem "Dulang"
    .AddItem "Design"
    .AddItem "Diamond"
    .AddItem "Lebar"
    .AddItem "Panjang"
    .AddItem "Purity"
    .AddItem "Upah Modal"
    .AddItem "Upah Jualan"
    .AddItem "Saiz"
    .AddItem "Supplier"
End With

With Frm56.CBB9
    .AddItem "No Data"
    .AddItem "Barcode"
    .AddItem "Berat"
    .AddItem "Berat Riyal"
    .AddItem "Code 1"
    .AddItem "Code 2"
    .AddItem "Dulang"
    .AddItem "Design"
    .AddItem "Diamond"
    .AddItem "Lebar"
    .AddItem "Panjang"
    .AddItem "Purity"
    .AddItem "Upah Modal"
    .AddItem "Upah Jualan"
    .AddItem "Saiz"
    .AddItem "Supplier"
End With

With Frm56.CBB14
    .AddItem "No Data"
    .AddItem "Barcode"
    .AddItem "Berat"
    .AddItem "Berat Riyal"
    .AddItem "Code 1"
    .AddItem "Code 2"
    .AddItem "Dulang"
    .AddItem "Design"
    .AddItem "Diamond"
    .AddItem "Lebar"
    .AddItem "Panjang"
    .AddItem "Purity"
    .AddItem "Upah Modal"
    .AddItem "Upah Jualan"
    .AddItem "Saiz"
    .AddItem "Supplier"
End With

With Frm56.CBB15
    .AddItem "No Data"
    .AddItem "Barcode"
    .AddItem "Berat"
    .AddItem "Berat Riyal"
    .AddItem "Code 1"
    .AddItem "Code 2"
    .AddItem "Dulang"
    .AddItem "Design"
    .AddItem "Diamond"
    .AddItem "Lebar"
    .AddItem "Panjang"
    .AddItem "Purity"
    .AddItem "Upah Modal"
    .AddItem "Upah Jualan"
    .AddItem "Saiz"
    .AddItem "Supplier"
End With

With Frm56.CBB16
    .AddItem "No Data"
    .AddItem "Barcode"
    .AddItem "Berat"
    .AddItem "Berat Riyal"
    .AddItem "Code 1"
    .AddItem "Code 2"
    .AddItem "Dulang"
    .AddItem "Design"
    .AddItem "Diamond"
    .AddItem "Lebar"
    .AddItem "Panjang"
    .AddItem "Purity"
    .AddItem "Upah Modal"
    .AddItem "Upah Jualan"
    .AddItem "Saiz"
    .AddItem "Supplier"
End With

Dim Frm56_LM_BARCODE(13)
Dim Frm56_LM_Jenis(13)

Frm56.CBB10.Clear
Frm56.CBB11.Clear
Frm56.CBB12.Clear
Frm56.CBB13.Clear

For i = 5 To 10

    Frm56.CBB10.AddItem i
    Frm56.CBB11.AddItem i
    Frm56.CBB12.AddItem i
    Frm56.CBB13.AddItem i
    
Next i

Frm56.CB1.Value = Split(G_SKU_LINE(0), ",")(0)
Frm56.CB2.Value = Split(G_SKU_LINE(0), ",")(1)
Frm56.CB3.Value = Split(G_SKU_LINE(0), ",")(2)
Frm56.CB4.Value = Split(G_SKU_LINE(0), ",")(3)
 
If Split(G_SKU_PRE_DATA(0), ",")(0) <> "" Then Frm56.CBB1 = Split(G_SKU_PRE_DATA(0), ",")(0)
If Split(G_SKU_PRE_DATA(0), ",")(1) <> "" Then Frm56.CBB2 = Split(G_SKU_PRE_DATA(0), ",")(1)
If Split(G_SKU_PRE_DATA(0), ",")(2) <> "" Then Frm56.CBB14 = Split(G_SKU_PRE_DATA(0), ",")(2)
If Split(G_SKU_PRE_DATA(0), ",")(3) <> "" Then Frm56.CBB3 = Split(G_SKU_PRE_DATA(0), ",")(3)
If Split(G_SKU_PRE_DATA(0), ",")(4) <> "" Then Frm56.CBB4 = Split(G_SKU_PRE_DATA(0), ",")(4)
If Split(G_SKU_PRE_DATA(0), ",")(5) <> "" Then Frm56.CBB5 = Split(G_SKU_PRE_DATA(0), ",")(5)
If Split(G_SKU_PRE_DATA(0), ",")(6) <> "" Then Frm56.CBB6 = Split(G_SKU_PRE_DATA(0), ",")(6)
If Split(G_SKU_PRE_DATA(0), ",")(7) <> "" Then Frm56.CBB7 = Split(G_SKU_PRE_DATA(0), ",")(7)
If Split(G_SKU_PRE_DATA(0), ",")(8) <> "" Then Frm56.CBB15 = Split(G_SKU_PRE_DATA(0), ",")(8)
If Split(G_SKU_PRE_DATA(0), ",")(9) <> "" Then Frm56.CBB8 = Split(G_SKU_PRE_DATA(0), ",")(9)
If Split(G_SKU_PRE_DATA(0), ",")(10) <> "" Then Frm56.CBB9 = Split(G_SKU_PRE_DATA(0), ",")(10)
If Split(G_SKU_PRE_DATA(0), ",")(11) <> "" Then Frm56.CBB16 = Split(G_SKU_PRE_DATA(0), ",")(11)

If Split(G_SKU_DATA(0), ",")(0) <> "No Data" Then Frm56.L3_Text = Split(G_SKU_DATA(0), ",")(0)
If Split(G_SKU_DATA(0), ",")(1) <> "No Data" Then Frm56.L4_Text = Split(G_SKU_DATA(0), ",")(1)
If Split(G_SKU_DATA(0), ",")(2) <> "No Data" Then Frm56.L12_Text = Split(G_SKU_DATA(0), ",")(2)
If Split(G_SKU_DATA(0), ",")(3) <> "No Data" Then Frm56.L5_Text = Split(G_SKU_DATA(0), ",")(3)
If Split(G_SKU_DATA(0), ",")(4) <> "No Data" Then Frm56.L6_Text = Split(G_SKU_DATA(0), ",")(4)
If Split(G_SKU_DATA(0), ",")(5) <> "No Data" Then Frm56.L7_Text = Split(G_SKU_DATA(0), ",")(5)
If Split(G_SKU_DATA(0), ",")(6) <> "No Data" Then Frm56.L8_Text = Split(G_SKU_DATA(0), ",")(6)
If Split(G_SKU_DATA(0), ",")(7) <> "No Data" Then Frm56.L9_Text = Split(G_SKU_DATA(0), ",")(7)
If Split(G_SKU_DATA(0), ",")(8) <> "No Data" Then Frm56.L13_Text = Split(G_SKU_DATA(0), ",")(8)
If Split(G_SKU_DATA(0), ",")(9) <> "No Data" Then Frm56.L10_Text = Split(G_SKU_DATA(0), ",")(9)
If Split(G_SKU_DATA(0), ",")(10) <> "No Data" Then Frm56.L11_Text = Split(G_SKU_DATA(0), ",")(10)
If Split(G_SKU_DATA(0), ",")(11) <> "No Data" Then Frm56.L14_Text = Split(G_SKU_DATA(0), ",")(11)

Exit Sub

logging:

LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " Frm56 : Frm56_ListItem_BK" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
Call log_rekod

Resume Next
End Sub
