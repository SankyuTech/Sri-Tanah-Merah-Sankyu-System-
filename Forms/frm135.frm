VERSION 5.00
Begin VB.Form frm135 
   Caption         =   "Printer Setting"
   ClientHeight    =   12645
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   21645
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
   MDIChild        =   -1  'True
   ScaleHeight     =   12645
   ScaleWidth      =   21645
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame46 
      Caption         =   "X3"
      Height          =   2655
      Left            =   15720
      TabIndex        =   225
      Top             =   9600
      Width           =   2295
      Begin VB.Frame Frame48 
         Caption         =   "Italic"
         Height          =   615
         Left            =   120
         TabIndex        =   233
         Top             =   1920
         Width           =   2055
         Begin VB.OptionButton Option64 
            Caption         =   "Ya"
            Height          =   255
            Left            =   240
            TabIndex        =   235
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton Option63 
            Caption         =   "Tidak"
            Height          =   255
            Left            =   1080
            TabIndex        =   234
            Top             =   240
            Width           =   790
         End
      End
      Begin VB.Frame Frame47 
         Caption         =   "Bold"
         Height          =   615
         Left            =   120
         TabIndex        =   230
         Top             =   1320
         Width           =   2055
         Begin VB.OptionButton Option62 
            Caption         =   "Tidak"
            Height          =   255
            Left            =   1080
            TabIndex        =   232
            Top             =   240
            Width           =   790
         End
         Begin VB.OptionButton Option61 
            Caption         =   "Ya"
            Height          =   255
            Left            =   240
            TabIndex        =   231
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.TextBox Text63 
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
         Height          =   285
         Left            =   900
         TabIndex        =   229
         Top             =   960
         Width           =   1000
      End
      Begin VB.TextBox Text62 
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
         Height          =   285
         Left            =   900
         TabIndex        =   228
         Top             =   720
         Width           =   1000
      End
      Begin VB.TextBox Text61 
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
         Height          =   285
         Left            =   900
         TabIndex        =   227
         Top             =   480
         Width           =   1000
      End
      Begin VB.TextBox Text60 
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
         Height          =   285
         Left            =   900
         TabIndex        =   226
         Top             =   240
         Width           =   1000
      End
      Begin VB.Label Label64 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Height :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   239
         Top             =   960
         Width           =   600
      End
      Begin VB.Label Label63 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Width :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   238
         Top             =   720
         Width           =   600
      End
      Begin VB.Label Label62 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "X :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   237
         Top             =   480
         Width           =   600
      End
      Begin VB.Label Label61 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "X :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   236
         Top             =   240
         Width           =   600
      End
   End
   Begin VB.Frame Frame43 
      Caption         =   "X3"
      Height          =   2655
      Left            =   18120
      TabIndex        =   210
      Top             =   8520
      Width           =   2295
      Begin VB.Frame Frame45 
         Caption         =   "Italic"
         Height          =   615
         Left            =   120
         TabIndex        =   218
         Top             =   1920
         Width           =   2055
         Begin VB.OptionButton Option60 
            Caption         =   "Ya"
            Height          =   255
            Left            =   240
            TabIndex        =   220
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton Option59 
            Caption         =   "Tidak"
            Height          =   255
            Left            =   1080
            TabIndex        =   219
            Top             =   240
            Width           =   790
         End
      End
      Begin VB.Frame Frame44 
         Caption         =   "Bold"
         Height          =   615
         Left            =   120
         TabIndex        =   215
         Top             =   1320
         Width           =   2055
         Begin VB.OptionButton Option58 
            Caption         =   "Tidak"
            Height          =   255
            Left            =   1080
            TabIndex        =   217
            Top             =   240
            Width           =   790
         End
         Begin VB.OptionButton Option57 
            Caption         =   "Ya"
            Height          =   255
            Left            =   240
            TabIndex        =   216
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.TextBox Text59 
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
         Height          =   285
         Left            =   900
         TabIndex        =   214
         Top             =   960
         Width           =   1000
      End
      Begin VB.TextBox Text58 
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
         Height          =   285
         Left            =   900
         TabIndex        =   213
         Top             =   720
         Width           =   1000
      End
      Begin VB.TextBox Text57 
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
         Height          =   285
         Left            =   900
         TabIndex        =   212
         Top             =   480
         Width           =   1000
      End
      Begin VB.TextBox Text56 
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
         Height          =   285
         Left            =   900
         TabIndex        =   211
         Top             =   240
         Width           =   1000
      End
      Begin VB.Label Label60 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Height :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   224
         Top             =   960
         Width           =   600
      End
      Begin VB.Label Label59 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Width :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   223
         Top             =   720
         Width           =   600
      End
      Begin VB.Label Label58 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "X :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   222
         Top             =   480
         Width           =   600
      End
      Begin VB.Label Label57 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "X :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   221
         Top             =   240
         Width           =   600
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "X3"
      Height          =   2655
      Left            =   120
      TabIndex        =   195
      Top             =   8400
      Width           =   2295
      Begin VB.Frame Frame6 
         Caption         =   "Italic"
         Height          =   615
         Left            =   120
         TabIndex        =   203
         Top             =   1920
         Width           =   2055
         Begin VB.OptionButton Option8 
            Caption         =   "Ya"
            Height          =   255
            Left            =   240
            TabIndex        =   205
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton Option7 
            Caption         =   "Tidak"
            Height          =   255
            Left            =   1080
            TabIndex        =   204
            Top             =   240
            Width           =   790
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Bold"
         Height          =   615
         Left            =   120
         TabIndex        =   200
         Top             =   1320
         Width           =   2055
         Begin VB.OptionButton Option6 
            Caption         =   "Tidak"
            Height          =   255
            Left            =   1080
            TabIndex        =   202
            Top             =   240
            Width           =   790
         End
         Begin VB.OptionButton Option5 
            Caption         =   "Ya"
            Height          =   255
            Left            =   240
            TabIndex        =   201
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.TextBox Text7 
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
         Height          =   285
         Left            =   900
         TabIndex        =   199
         Top             =   960
         Width           =   1000
      End
      Begin VB.TextBox Text6 
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
         Height          =   285
         Left            =   900
         TabIndex        =   198
         Top             =   720
         Width           =   1000
      End
      Begin VB.TextBox Text5 
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
         Height          =   285
         Left            =   900
         TabIndex        =   197
         Top             =   480
         Width           =   1000
      End
      Begin VB.TextBox Text4 
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
         Height          =   285
         Left            =   900
         TabIndex        =   196
         Top             =   240
         Width           =   1000
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Height :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   209
         Top             =   960
         Width           =   600
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Width :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   208
         Top             =   720
         Width           =   600
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "X :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   207
         Top             =   480
         Width           =   600
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "X :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   206
         Top             =   240
         Width           =   600
      End
   End
   Begin VB.Frame Frame40 
      Caption         =   "n"
      Height          =   2655
      Left            =   10440
      TabIndex        =   180
      Top             =   9840
      Width           =   2295
      Begin VB.Frame Frame42 
         Caption         =   "Italic"
         Height          =   615
         Left            =   120
         TabIndex        =   188
         Top             =   1920
         Width           =   2055
         Begin VB.OptionButton Option56 
            Caption         =   "Ya"
            Height          =   255
            Left            =   240
            TabIndex        =   190
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton Option55 
            Caption         =   "Tidak"
            Height          =   255
            Left            =   1080
            TabIndex        =   189
            Top             =   240
            Width           =   790
         End
      End
      Begin VB.Frame Frame41 
         Caption         =   "Bold"
         Height          =   615
         Left            =   120
         TabIndex        =   185
         Top             =   1320
         Width           =   2055
         Begin VB.OptionButton Option54 
            Caption         =   "Tidak"
            Height          =   255
            Left            =   1080
            TabIndex        =   187
            Top             =   240
            Width           =   790
         End
         Begin VB.OptionButton Option53 
            Caption         =   "Ya"
            Height          =   255
            Left            =   240
            TabIndex        =   186
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.TextBox Text55 
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
         Height          =   285
         Left            =   900
         TabIndex        =   184
         Top             =   960
         Width           =   1000
      End
      Begin VB.TextBox Text54 
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
         Height          =   285
         Left            =   900
         TabIndex        =   183
         Top             =   720
         Width           =   1000
      End
      Begin VB.TextBox Text53 
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
         Height          =   285
         Left            =   900
         TabIndex        =   182
         Top             =   480
         Width           =   1000
      End
      Begin VB.TextBox Text52 
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
         Height          =   285
         Left            =   900
         TabIndex        =   181
         Top             =   240
         Width           =   1000
      End
      Begin VB.Label Label56 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Height :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   194
         Top             =   960
         Width           =   600
      End
      Begin VB.Label Label55 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Width :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   193
         Top             =   720
         Width           =   600
      End
      Begin VB.Label Label54 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "X :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   192
         Top             =   480
         Width           =   600
      End
      Begin VB.Label Label53 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "X :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   191
         Top             =   240
         Width           =   600
      End
   End
   Begin VB.Frame Frame37 
      Caption         =   "n"
      Height          =   2655
      Left            =   8400
      TabIndex        =   165
      Top             =   6360
      Width           =   2295
      Begin VB.Frame Frame39 
         Caption         =   "Italic"
         Height          =   615
         Left            =   120
         TabIndex        =   173
         Top             =   1920
         Width           =   2055
         Begin VB.OptionButton Option52 
            Caption         =   "Ya"
            Height          =   255
            Left            =   240
            TabIndex        =   175
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton Option51 
            Caption         =   "Tidak"
            Height          =   255
            Left            =   1080
            TabIndex        =   174
            Top             =   240
            Width           =   790
         End
      End
      Begin VB.Frame Frame38 
         Caption         =   "Bold"
         Height          =   615
         Left            =   120
         TabIndex        =   170
         Top             =   1320
         Width           =   2055
         Begin VB.OptionButton Option50 
            Caption         =   "Tidak"
            Height          =   255
            Left            =   1080
            TabIndex        =   172
            Top             =   240
            Width           =   790
         End
         Begin VB.OptionButton Option49 
            Caption         =   "Ya"
            Height          =   255
            Left            =   240
            TabIndex        =   171
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.TextBox Text51 
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
         Height          =   285
         Left            =   900
         TabIndex        =   169
         Top             =   960
         Width           =   1000
      End
      Begin VB.TextBox Text50 
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
         Height          =   285
         Left            =   900
         TabIndex        =   168
         Top             =   720
         Width           =   1000
      End
      Begin VB.TextBox Text49 
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
         Height          =   285
         Left            =   900
         TabIndex        =   167
         Top             =   480
         Width           =   1000
      End
      Begin VB.TextBox Text48 
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
         Height          =   285
         Left            =   900
         TabIndex        =   166
         Top             =   240
         Width           =   1000
      End
      Begin VB.Label Label52 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Height :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   179
         Top             =   960
         Width           =   600
      End
      Begin VB.Label Label51 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Width :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   178
         Top             =   720
         Width           =   600
      End
      Begin VB.Label Label50 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "X :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   177
         Top             =   480
         Width           =   600
      End
      Begin VB.Label Label49 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "X :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   176
         Top             =   240
         Width           =   600
      End
   End
   Begin VB.Frame Frame34 
      Caption         =   "n"
      Height          =   2655
      Left            =   10800
      TabIndex        =   150
      Top             =   6360
      Width           =   2295
      Begin VB.Frame Frame36 
         Caption         =   "Italic"
         Height          =   615
         Left            =   120
         TabIndex        =   158
         Top             =   1920
         Width           =   2055
         Begin VB.OptionButton Option48 
            Caption         =   "Ya"
            Height          =   255
            Left            =   240
            TabIndex        =   160
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton Option47 
            Caption         =   "Tidak"
            Height          =   255
            Left            =   1080
            TabIndex        =   159
            Top             =   240
            Width           =   790
         End
      End
      Begin VB.Frame Frame35 
         Caption         =   "Bold"
         Height          =   615
         Left            =   120
         TabIndex        =   155
         Top             =   1320
         Width           =   2055
         Begin VB.OptionButton Option46 
            Caption         =   "Tidak"
            Height          =   255
            Left            =   1080
            TabIndex        =   157
            Top             =   240
            Width           =   790
         End
         Begin VB.OptionButton Option45 
            Caption         =   "Ya"
            Height          =   255
            Left            =   240
            TabIndex        =   156
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.TextBox Text47 
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
         Height          =   285
         Left            =   900
         TabIndex        =   154
         Top             =   960
         Width           =   1000
      End
      Begin VB.TextBox Text46 
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
         Height          =   285
         Left            =   900
         TabIndex        =   153
         Top             =   720
         Width           =   1000
      End
      Begin VB.TextBox Text45 
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
         Height          =   285
         Left            =   900
         TabIndex        =   152
         Top             =   480
         Width           =   1000
      End
      Begin VB.TextBox Text44 
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
         Height          =   285
         Left            =   900
         TabIndex        =   151
         Top             =   240
         Width           =   1000
      End
      Begin VB.Label Label48 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Height :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   164
         Top             =   960
         Width           =   600
      End
      Begin VB.Label Label47 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Width :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   163
         Top             =   720
         Width           =   600
      End
      Begin VB.Label Label46 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "X :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   162
         Top             =   480
         Width           =   600
      End
      Begin VB.Label Label45 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "X :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   161
         Top             =   240
         Width           =   600
      End
   End
   Begin VB.Frame Frame31 
      Caption         =   "n"
      Height          =   2655
      Left            =   13320
      TabIndex        =   135
      Top             =   9480
      Width           =   2295
      Begin VB.Frame Frame33 
         Caption         =   "Italic"
         Height          =   615
         Left            =   120
         TabIndex        =   143
         Top             =   1920
         Width           =   2055
         Begin VB.OptionButton Option44 
            Caption         =   "Ya"
            Height          =   255
            Left            =   240
            TabIndex        =   145
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton Option43 
            Caption         =   "Tidak"
            Height          =   255
            Left            =   1080
            TabIndex        =   144
            Top             =   240
            Width           =   790
         End
      End
      Begin VB.Frame Frame32 
         Caption         =   "Bold"
         Height          =   615
         Left            =   120
         TabIndex        =   140
         Top             =   1320
         Width           =   2055
         Begin VB.OptionButton Option42 
            Caption         =   "Tidak"
            Height          =   255
            Left            =   1080
            TabIndex        =   142
            Top             =   240
            Width           =   790
         End
         Begin VB.OptionButton Option41 
            Caption         =   "Ya"
            Height          =   255
            Left            =   240
            TabIndex        =   141
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.TextBox Text43 
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
         Height          =   285
         Left            =   900
         TabIndex        =   139
         Top             =   960
         Width           =   1000
      End
      Begin VB.TextBox Text42 
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
         Height          =   285
         Left            =   900
         TabIndex        =   138
         Top             =   720
         Width           =   1000
      End
      Begin VB.TextBox Text41 
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
         Height          =   285
         Left            =   900
         TabIndex        =   137
         Top             =   480
         Width           =   1000
      End
      Begin VB.TextBox Text40 
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
         Height          =   285
         Left            =   900
         TabIndex        =   136
         Top             =   240
         Width           =   1000
      End
      Begin VB.Label Label44 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Height :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   149
         Top             =   960
         Width           =   600
      End
      Begin VB.Label Label43 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Width :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   148
         Top             =   720
         Width           =   600
      End
      Begin VB.Label Label42 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "X :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   147
         Top             =   480
         Width           =   600
      End
      Begin VB.Label Label41 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "X :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   146
         Top             =   240
         Width           =   600
      End
   End
   Begin VB.Frame Frame28 
      Caption         =   "n"
      Height          =   2655
      Left            =   2520
      TabIndex        =   120
      Top             =   10080
      Width           =   2295
      Begin VB.Frame Frame30 
         Caption         =   "Italic"
         Height          =   615
         Left            =   120
         TabIndex        =   128
         Top             =   1920
         Width           =   2055
         Begin VB.OptionButton Option40 
            Caption         =   "Ya"
            Height          =   255
            Left            =   240
            TabIndex        =   130
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton Option39 
            Caption         =   "Tidak"
            Height          =   255
            Left            =   1080
            TabIndex        =   129
            Top             =   240
            Width           =   790
         End
      End
      Begin VB.Frame Frame29 
         Caption         =   "Bold"
         Height          =   615
         Left            =   120
         TabIndex        =   125
         Top             =   1320
         Width           =   2055
         Begin VB.OptionButton Option38 
            Caption         =   "Tidak"
            Height          =   255
            Left            =   1080
            TabIndex        =   127
            Top             =   240
            Width           =   790
         End
         Begin VB.OptionButton Option37 
            Caption         =   "Ya"
            Height          =   255
            Left            =   240
            TabIndex        =   126
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.TextBox Text39 
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
         Height          =   285
         Left            =   900
         TabIndex        =   124
         Top             =   960
         Width           =   1000
      End
      Begin VB.TextBox Text38 
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
         Height          =   285
         Left            =   900
         TabIndex        =   123
         Top             =   720
         Width           =   1000
      End
      Begin VB.TextBox Text37 
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
         Height          =   285
         Left            =   900
         TabIndex        =   122
         Top             =   480
         Width           =   1000
      End
      Begin VB.TextBox Text36 
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
         Height          =   285
         Left            =   900
         TabIndex        =   121
         Top             =   240
         Width           =   1000
      End
      Begin VB.Label Label40 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Height :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   134
         Top             =   960
         Width           =   600
      End
      Begin VB.Label Label39 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Width :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   133
         Top             =   720
         Width           =   600
      End
      Begin VB.Label Label38 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "X :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   132
         Top             =   480
         Width           =   600
      End
      Begin VB.Label Label37 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "X :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   131
         Top             =   240
         Width           =   600
      End
   End
   Begin VB.Frame Frame25 
      Caption         =   "X3"
      Height          =   2655
      Left            =   6360
      TabIndex        =   105
      Top             =   9960
      Width           =   2295
      Begin VB.Frame Frame27 
         Caption         =   "Italic"
         Height          =   615
         Left            =   120
         TabIndex        =   113
         Top             =   1920
         Width           =   2055
         Begin VB.OptionButton Option36 
            Caption         =   "Ya"
            Height          =   255
            Left            =   240
            TabIndex        =   115
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton Option35 
            Caption         =   "Tidak"
            Height          =   255
            Left            =   1080
            TabIndex        =   114
            Top             =   240
            Width           =   790
         End
      End
      Begin VB.Frame Frame26 
         Caption         =   "Bold"
         Height          =   615
         Left            =   120
         TabIndex        =   110
         Top             =   1320
         Width           =   2055
         Begin VB.OptionButton Option34 
            Caption         =   "Tidak"
            Height          =   255
            Left            =   1080
            TabIndex        =   112
            Top             =   240
            Width           =   790
         End
         Begin VB.OptionButton Option33 
            Caption         =   "Ya"
            Height          =   255
            Left            =   240
            TabIndex        =   111
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.TextBox Text35 
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
         Height          =   285
         Left            =   900
         TabIndex        =   109
         Top             =   960
         Width           =   1000
      End
      Begin VB.TextBox Text34 
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
         Height          =   285
         Left            =   900
         TabIndex        =   108
         Top             =   720
         Width           =   1000
      End
      Begin VB.TextBox Text33 
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
         Height          =   285
         Left            =   900
         TabIndex        =   107
         Top             =   480
         Width           =   1000
      End
      Begin VB.TextBox Text32 
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
         Height          =   285
         Left            =   900
         TabIndex        =   106
         Top             =   240
         Width           =   1000
      End
      Begin VB.Label Label36 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Height :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   119
         Top             =   960
         Width           =   600
      End
      Begin VB.Label Label35 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Width :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   118
         Top             =   720
         Width           =   600
      End
      Begin VB.Label Label34 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "X :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   117
         Top             =   480
         Width           =   600
      End
      Begin VB.Label Label33 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "X :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   116
         Top             =   240
         Width           =   600
      End
   End
   Begin VB.Frame Frame22 
      Caption         =   "X3"
      Height          =   2655
      Left            =   18120
      TabIndex        =   90
      Top             =   5760
      Width           =   2295
      Begin VB.Frame Frame24 
         Caption         =   "Italic"
         Height          =   615
         Left            =   120
         TabIndex        =   98
         Top             =   1920
         Width           =   2055
         Begin VB.OptionButton Option32 
            Caption         =   "Ya"
            Height          =   255
            Left            =   240
            TabIndex        =   100
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton Option31 
            Caption         =   "Tidak"
            Height          =   255
            Left            =   1080
            TabIndex        =   99
            Top             =   240
            Width           =   790
         End
      End
      Begin VB.Frame Frame23 
         Caption         =   "Bold"
         Height          =   615
         Left            =   120
         TabIndex        =   95
         Top             =   1320
         Width           =   2055
         Begin VB.OptionButton Option30 
            Caption         =   "Tidak"
            Height          =   255
            Left            =   1080
            TabIndex        =   97
            Top             =   240
            Width           =   790
         End
         Begin VB.OptionButton Option29 
            Caption         =   "Ya"
            Height          =   255
            Left            =   240
            TabIndex        =   96
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.TextBox Text31 
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
         Height          =   285
         Left            =   900
         TabIndex        =   94
         Top             =   960
         Width           =   1000
      End
      Begin VB.TextBox Text30 
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
         Height          =   285
         Left            =   900
         TabIndex        =   93
         Top             =   720
         Width           =   1000
      End
      Begin VB.TextBox Text29 
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
         Height          =   285
         Left            =   900
         TabIndex        =   92
         Top             =   480
         Width           =   1000
      End
      Begin VB.TextBox Text28 
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
         Height          =   285
         Left            =   900
         TabIndex        =   91
         Top             =   240
         Width           =   1000
      End
      Begin VB.Label Label32 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Height :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   104
         Top             =   960
         Width           =   600
      End
      Begin VB.Label Label31 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Width :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   103
         Top             =   720
         Width           =   600
      End
      Begin VB.Label Label30 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "X :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   102
         Top             =   480
         Width           =   600
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "X :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   101
         Top             =   240
         Width           =   600
      End
   End
   Begin VB.Frame Frame13 
      Caption         =   "X3"
      Height          =   2655
      Left            =   18120
      TabIndex        =   60
      Top             =   3000
      Width           =   2295
      Begin VB.Frame Frame19 
         Caption         =   "X3"
         Height          =   2655
         Left            =   2040
         TabIndex        =   75
         Top             =   2520
         Width           =   2295
         Begin VB.Frame Frame21 
            Caption         =   "Italic"
            Height          =   615
            Left            =   120
            TabIndex        =   83
            Top             =   1920
            Width           =   2055
            Begin VB.OptionButton Option28 
               Caption         =   "Ya"
               Height          =   255
               Left            =   240
               TabIndex        =   85
               Top             =   240
               Width           =   735
            End
            Begin VB.OptionButton Option27 
               Caption         =   "Tidak"
               Height          =   255
               Left            =   1080
               TabIndex        =   84
               Top             =   240
               Width           =   790
            End
         End
         Begin VB.Frame Frame20 
            Caption         =   "Bold"
            Height          =   615
            Left            =   120
            TabIndex        =   80
            Top             =   1320
            Width           =   2055
            Begin VB.OptionButton Option26 
               Caption         =   "Tidak"
               Height          =   255
               Left            =   1080
               TabIndex        =   82
               Top             =   240
               Width           =   790
            End
            Begin VB.OptionButton Option25 
               Caption         =   "Ya"
               Height          =   255
               Left            =   240
               TabIndex        =   81
               Top             =   240
               Width           =   735
            End
         End
         Begin VB.TextBox Text27 
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
            Height          =   285
            Left            =   900
            TabIndex        =   79
            Top             =   960
            Width           =   1000
         End
         Begin VB.TextBox Text26 
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
            Height          =   285
            Left            =   900
            TabIndex        =   78
            Top             =   720
            Width           =   1000
         End
         Begin VB.TextBox Text25 
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
            Height          =   285
            Left            =   900
            TabIndex        =   77
            Top             =   480
            Width           =   1000
         End
         Begin VB.TextBox Text24 
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
            Height          =   285
            Left            =   900
            TabIndex        =   76
            Top             =   240
            Width           =   1000
         End
         Begin VB.Label Label28 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Height :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   240
            TabIndex        =   89
            Top             =   960
            Width           =   600
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Width :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   240
            TabIndex        =   88
            Top             =   720
            Width           =   600
         End
         Begin VB.Label Label26 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "X :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   240
            TabIndex        =   87
            Top             =   480
            Width           =   600
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "X :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   240
            TabIndex        =   86
            Top             =   240
            Width           =   600
         End
      End
      Begin VB.Frame Frame14 
         Caption         =   "Italic"
         Height          =   615
         Left            =   120
         TabIndex        =   68
         Top             =   1920
         Width           =   2055
         Begin VB.OptionButton Option17 
            Caption         =   "Ya"
            Height          =   255
            Left            =   240
            TabIndex        =   70
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton Option18 
            Caption         =   "Tidak"
            Height          =   255
            Left            =   1080
            TabIndex        =   69
            Top             =   240
            Width           =   790
         End
      End
      Begin VB.Frame Frame15 
         Caption         =   "Bold"
         Height          =   615
         Left            =   120
         TabIndex        =   65
         Top             =   1320
         Width           =   2055
         Begin VB.OptionButton Option19 
            Caption         =   "Tidak"
            Height          =   255
            Left            =   1080
            TabIndex        =   67
            Top             =   240
            Width           =   790
         End
         Begin VB.OptionButton Option20 
            Caption         =   "Ya"
            Height          =   255
            Left            =   240
            TabIndex        =   66
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.TextBox Text16 
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
         Height          =   285
         Left            =   900
         TabIndex        =   64
         Top             =   960
         Width           =   1000
      End
      Begin VB.TextBox Text17 
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
         Height          =   285
         Left            =   900
         TabIndex        =   63
         Top             =   720
         Width           =   1000
      End
      Begin VB.TextBox Text18 
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
         Height          =   285
         Left            =   900
         TabIndex        =   62
         Top             =   480
         Width           =   1000
      End
      Begin VB.TextBox Text19 
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
         Height          =   285
         Left            =   900
         TabIndex        =   61
         Top             =   240
         Width           =   1000
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Height :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   74
         Top             =   960
         Width           =   600
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Width :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   73
         Top             =   720
         Width           =   600
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "X :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   72
         Top             =   480
         Width           =   600
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "X :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   71
         Top             =   240
         Width           =   600
      End
   End
   Begin VB.Frame Frame16 
      Caption         =   "X3"
      Height          =   2655
      Left            =   18120
      TabIndex        =   45
      Top             =   240
      Width           =   2295
      Begin VB.Frame Frame18 
         Caption         =   "Italic"
         Height          =   615
         Left            =   120
         TabIndex        =   53
         Top             =   1920
         Width           =   2055
         Begin VB.OptionButton Option24 
            Caption         =   "Ya"
            Height          =   255
            Left            =   240
            TabIndex        =   55
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton Option23 
            Caption         =   "Tidak"
            Height          =   255
            Left            =   1080
            TabIndex        =   54
            Top             =   240
            Width           =   790
         End
      End
      Begin VB.Frame Frame17 
         Caption         =   "Bold"
         Height          =   615
         Left            =   120
         TabIndex        =   50
         Top             =   1320
         Width           =   2055
         Begin VB.OptionButton Option22 
            Caption         =   "Tidak"
            Height          =   255
            Left            =   1080
            TabIndex        =   52
            Top             =   240
            Width           =   790
         End
         Begin VB.OptionButton Option21 
            Caption         =   "Ya"
            Height          =   255
            Left            =   240
            TabIndex        =   51
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.TextBox Text23 
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
         Height          =   285
         Left            =   900
         TabIndex        =   49
         Top             =   960
         Width           =   1000
      End
      Begin VB.TextBox Text22 
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
         Height          =   285
         Left            =   900
         TabIndex        =   48
         Top             =   720
         Width           =   1000
      End
      Begin VB.TextBox Text21 
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
         Height          =   285
         Left            =   900
         TabIndex        =   47
         Top             =   480
         Width           =   1000
      End
      Begin VB.TextBox Text20 
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
         Height          =   285
         Left            =   900
         TabIndex        =   46
         Top             =   240
         Width           =   1000
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Height :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   59
         Top             =   960
         Width           =   600
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Width :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   58
         Top             =   720
         Width           =   600
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "X :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   57
         Top             =   480
         Width           =   600
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "X :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   56
         Top             =   240
         Width           =   600
      End
   End
   Begin VB.Frame Frame10 
      Caption         =   "X3"
      Height          =   2655
      Left            =   120
      TabIndex        =   30
      Top             =   5640
      Width           =   2295
      Begin VB.Frame Frame12 
         Caption         =   "Italic"
         Height          =   615
         Left            =   120
         TabIndex        =   38
         Top             =   1920
         Width           =   2055
         Begin VB.OptionButton Option16 
            Caption         =   "Ya"
            Height          =   255
            Left            =   240
            TabIndex        =   40
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton Option15 
            Caption         =   "Tidak"
            Height          =   255
            Left            =   1080
            TabIndex        =   39
            Top             =   240
            Width           =   790
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Bold"
         Height          =   615
         Left            =   120
         TabIndex        =   35
         Top             =   1320
         Width           =   2055
         Begin VB.OptionButton Option14 
            Caption         =   "Tidak"
            Height          =   255
            Left            =   1080
            TabIndex        =   37
            Top             =   240
            Width           =   790
         End
         Begin VB.OptionButton Option13 
            Caption         =   "Ya"
            Height          =   255
            Left            =   240
            TabIndex        =   36
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.TextBox Text15 
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
         Height          =   285
         Left            =   900
         TabIndex        =   34
         Top             =   960
         Width           =   1000
      End
      Begin VB.TextBox Text14 
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
         Height          =   285
         Left            =   900
         TabIndex        =   33
         Top             =   720
         Width           =   1000
      End
      Begin VB.TextBox Text13 
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
         Height          =   285
         Left            =   900
         TabIndex        =   32
         Top             =   480
         Width           =   1000
      End
      Begin VB.TextBox Text12 
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
         Height          =   285
         Left            =   900
         TabIndex        =   31
         Top             =   240
         Width           =   1000
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Height :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   44
         Top             =   960
         Width           =   600
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Width :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   43
         Top             =   720
         Width           =   600
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "X :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   42
         Top             =   480
         Width           =   600
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "X :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   41
         Top             =   240
         Width           =   600
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "X2"
      Height          =   2655
      Left            =   120
      TabIndex        =   15
      Top             =   2880
      Width           =   2295
      Begin VB.Frame Frame9 
         Caption         =   "Italic"
         Height          =   615
         Left            =   120
         TabIndex        =   23
         Top             =   1920
         Width           =   2055
         Begin VB.OptionButton Option12 
            Caption         =   "Ya"
            Height          =   255
            Left            =   240
            TabIndex        =   25
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton Option11 
            Caption         =   "Tidak"
            Height          =   255
            Left            =   1080
            TabIndex        =   24
            Top             =   240
            Width           =   790
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Bold"
         Height          =   615
         Left            =   120
         TabIndex        =   20
         Top             =   1320
         Width           =   2055
         Begin VB.OptionButton Option10 
            Caption         =   "Tidak"
            Height          =   255
            Left            =   1080
            TabIndex        =   22
            Top             =   240
            Width           =   790
         End
         Begin VB.OptionButton Option9 
            Caption         =   "Ya"
            Height          =   255
            Left            =   240
            TabIndex        =   21
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.TextBox Text11 
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
         Height          =   285
         Left            =   900
         TabIndex        =   19
         Top             =   960
         Width           =   1000
      End
      Begin VB.TextBox Text10 
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
         Height          =   285
         Left            =   900
         TabIndex        =   18
         Top             =   720
         Width           =   1000
      End
      Begin VB.TextBox Text9 
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
         Height          =   285
         Left            =   900
         TabIndex        =   17
         Top             =   480
         Width           =   1000
      End
      Begin VB.TextBox Text8 
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
         Height          =   285
         Left            =   900
         TabIndex        =   16
         Top             =   240
         Width           =   1000
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Height :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   29
         Top             =   960
         Width           =   600
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Width :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   28
         Top             =   720
         Width           =   600
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "X :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   27
         Top             =   480
         Width           =   600
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "X :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   26
         Top             =   240
         Width           =   600
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "X1"
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
      Begin VB.Frame Frame3 
         Caption         =   "Italic"
         Height          =   615
         Left            =   120
         TabIndex        =   12
         Top             =   1920
         Width           =   2055
         Begin VB.OptionButton Option4 
            Caption         =   "Ya"
            Height          =   255
            Left            =   240
            TabIndex        =   14
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Tidak"
            Height          =   255
            Left            =   1080
            TabIndex        =   13
            Top             =   240
            Width           =   790
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Bold"
         Height          =   615
         Left            =   120
         TabIndex        =   9
         Top             =   1320
         Width           =   2055
         Begin VB.OptionButton Option2 
            Caption         =   "Tidak"
            Height          =   255
            Left            =   1080
            TabIndex        =   11
            Top             =   240
            Width           =   790
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Ya"
            Height          =   255
            Left            =   240
            TabIndex        =   10
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.TextBox Text3 
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
         Height          =   285
         Left            =   900
         TabIndex        =   7
         Top             =   960
         Width           =   1000
      End
      Begin VB.TextBox Text2 
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
         Height          =   285
         Left            =   900
         TabIndex        =   5
         Top             =   720
         Width           =   1000
      End
      Begin VB.TextBox Text1 
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
         Height          =   285
         Left            =   900
         TabIndex        =   3
         Top             =   480
         Width           =   1000
      End
      Begin VB.TextBox TB2 
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
         Height          =   285
         Left            =   900
         TabIndex        =   1
         Top             =   240
         Width           =   1000
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Height :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   8
         Top             =   960
         Width           =   600
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Width :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   600
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "X :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   600
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "X :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   600
      End
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   2520
      X2              =   4320
      Y1              =   6000
      Y2              =   3120
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   2520
      X2              =   4200
      Y1              =   3840
      Y2              =   2880
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   2520
      X2              =   4200
      Y1              =   1080
      Y2              =   2520
   End
   Begin VB.Image Image1 
      Height          =   10860
      Left            =   3120
      Picture         =   "frm135.frx":0000
      Top             =   720
      Width           =   14805
   End
End
Attribute VB_Name = "frm135"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
