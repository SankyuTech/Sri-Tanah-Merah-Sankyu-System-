VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm102 
   Caption         =   "Jualan kepada agen"
   ClientHeight    =   13035
   ClientLeft      =   120
   ClientTop       =   -20550
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
   Icon            =   "Frm102.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   13035
   ScaleWidth      =   23760
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Pic1 
      BorderStyle     =   0  'None
      Height          =   3255
      Left            =   0
      ScaleHeight     =   3255
      ScaleWidth      =   7815
      TabIndex        =   190
      Top             =   6360
      Width           =   7815
      Begin VB.CommandButton CMD13 
         Caption         =   "Tutup paparan ini"
         Height          =   360
         Left            =   1680
         MouseIcon       =   "Frm102.frx":0ECA
         MousePointer    =   99  'Custom
         TabIndex        =   206
         Top             =   2160
         Width           =   2295
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Maklumat cukai GST"
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
         TabIndex        =   205
         Top             =   240
         Width           =   5385
      End
      Begin VB.Label Label114 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Harga Tanpa GST"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   204
         Top             =   600
         Width           =   2505
      End
      Begin VB.Label Label117 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Harga Dengan GST"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   203
         Top             =   840
         Width           =   2505
      End
      Begin VB.Label Label111 
         BackStyle       =   0  'Transparent
         Caption         =   ": RM"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2640
         TabIndex        =   202
         Top             =   600
         Width           =   600
      End
      Begin VB.Label L15_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L15_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3075
         TabIndex        =   201
         Top             =   600
         Width           =   1545
      End
      Begin VB.Label Label115 
         BackStyle       =   0  'Transparent
         Caption         =   ": RM"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2640
         TabIndex        =   200
         Top             =   840
         Width           =   600
      End
      Begin VB.Label L16_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L16_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3075
         TabIndex        =   199
         Top             =   840
         Width           =   1545
      End
      Begin VB.Label L17_Text 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "L17_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2280
         TabIndex        =   198
         Top             =   1440
         Width           =   1785
      End
      Begin VB.Label Label121 
         BackStyle       =   0  'Transparent
         Caption         =   "Zero Rated (ZR)"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   197
         Top             =   1440
         Width           =   2145
      End
      Begin VB.Label Label118 
         BackStyle       =   0  'Transparent
         Caption         =   "Standard Rated (SR)"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   196
         Top             =   1680
         Width           =   2145
      End
      Begin VB.Label Label122 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Harga (RM)"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2280
         TabIndex        =   195
         Top             =   1200
         Width           =   1800
      End
      Begin VB.Label Label123 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah GST (RM)"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3960
         TabIndex        =   194
         Top             =   1200
         Width           =   1800
      End
      Begin VB.Label L18_Text 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "L18_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2280
         TabIndex        =   193
         Top             =   1680
         Width           =   1785
      End
      Begin VB.Label L19_Text 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "L19_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3960
         TabIndex        =   192
         Top             =   1440
         Width           =   1785
      End
      Begin VB.Label L20_Text 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "L20_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3960
         TabIndex        =   191
         Top             =   1680
         Width           =   1785
      End
      Begin VB.Line Line1 
         X1              =   2355
         X2              =   5635
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Shape Shape4 
         Height          =   1575
         Left            =   120
         Top             =   480
         Width           =   5655
      End
   End
   Begin VB.CommandButton CDM13 
      Caption         =   "Papar Maklumat Terperinci GST"
      Height          =   360
      Left            =   4800
      MouseIcon       =   "Frm102.frx":11D4
      MousePointer    =   99  'Custom
      TabIndex        =   207
      Top             =   8880
      Width           =   2895
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C0FF&
      Height          =   2175
      Left            =   8160
      ScaleHeight     =   2115
      ScaleWidth      =   5835
      TabIndex        =   167
      Top             =   360
      Visible         =   0   'False
      Width           =   5895
      Begin VB.Label L22_Text 
         Caption         =   "L22_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   0
         TabIndex        =   189
         Top             =   240
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label L23_Text 
         Caption         =   "L23_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   0
         TabIndex        =   188
         Top             =   480
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label L24_Text 
         Caption         =   "L24_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   0
         TabIndex        =   187
         Top             =   720
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label L25_Text 
         Caption         =   "L25_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   0
         TabIndex        =   186
         Top             =   960
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label L21_Text 
         Caption         =   "L21_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   0
         TabIndex        =   185
         Top             =   0
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label L29_Text 
         Caption         =   "L29_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   0
         TabIndex        =   184
         Top             =   1200
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label L31_Text 
         Caption         =   "L31_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1080
         TabIndex        =   183
         Top             =   240
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label L30_Text 
         Caption         =   "L30_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1080
         TabIndex        =   182
         Top             =   0
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label L32_Text 
         Caption         =   "L32_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1080
         TabIndex        =   181
         Top             =   480
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label L33_Text 
         Caption         =   "L33_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1080
         TabIndex        =   180
         Top             =   720
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label L40_Text 
         Caption         =   "L40_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2160
         TabIndex        =   179
         Top             =   1200
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label L35_Text 
         Caption         =   "L35_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2160
         TabIndex        =   178
         Top             =   0
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label L39_Text 
         Caption         =   "L39_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2160
         TabIndex        =   177
         Top             =   960
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label L38_Text 
         Caption         =   "L38_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2160
         TabIndex        =   176
         Top             =   720
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label L37_Text 
         Caption         =   "L37_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2160
         TabIndex        =   175
         Top             =   480
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label L36_Text 
         Caption         =   "L36_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2160
         TabIndex        =   174
         Top             =   240
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label L41_Text 
         Caption         =   "L41_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2160
         TabIndex        =   173
         Top             =   1440
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label L42_Text 
         Caption         =   "L42_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2160
         TabIndex        =   172
         Top             =   1680
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label L45_Text 
         Caption         =   "L45_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3240
         TabIndex        =   171
         Top             =   0
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label L47_Text 
         Caption         =   "L47_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3240
         TabIndex        =   170
         Top             =   360
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label L49_Text 
         Caption         =   "L49_Text"
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
         Left            =   3240
         TabIndex        =   169
         Top             =   720
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label L50_Text 
         Caption         =   "L50_Text"
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
         Left            =   3240
         TabIndex        =   168
         Top             =   1080
         Visible         =   0   'False
         Width           =   645
      End
   End
   Begin VB.CommandButton CMD12 
      Caption         =   "Maklumat Agen"
      Height          =   360
      Left            =   8280
      MouseIcon       =   "Frm102.frx":14DE
      MousePointer    =   99  'Custom
      TabIndex        =   166
      ToolTipText     =   "Sila pastikan semua data telah dimasukkan dengan betul sebelum masukkan data ini ke dalam senarai jualan."
      Top             =   10080
      Width           =   2055
   End
   Begin VB.CommandButton CMD11 
      Caption         =   "Batal"
      Height          =   360
      Left            =   9360
      MouseIcon       =   "Frm102.frx":17E8
      MousePointer    =   99  'Custom
      TabIndex        =   165
      ToolTipText     =   "Sila pastikan semua data telah dimasukkan dengan betul sebelum masukkan data ini ke dalam senarai jualan."
      Top             =   11040
      Width           =   1935
   End
   Begin VB.CommandButton CMD9 
      Caption         =   "Keluar"
      Height          =   360
      Left            =   9360
      MouseIcon       =   "Frm102.frx":1AF2
      MousePointer    =   99  'Custom
      TabIndex        =   164
      ToolTipText     =   "Sila pastikan semua data telah dimasukkan dengan betul sebelum masukkan data ini ke dalam senarai jualan."
      Top             =   11040
      Width           =   1935
   End
   Begin VB.CommandButton CMD10 
      Caption         =   "Jualan"
      Height          =   360
      Left            =   7320
      MouseIcon       =   "Frm102.frx":1DFC
      MousePointer    =   99  'Custom
      TabIndex        =   163
      ToolTipText     =   "Sila pastikan semua data telah dimasukkan dengan betul sebelum masukkan data ini ke dalam senarai jualan."
      Top             =   11040
      Width           =   1935
   End
   Begin VB.CommandButton CMD8 
      Caption         =   "Jualan"
      Height          =   360
      Left            =   7320
      MouseIcon       =   "Frm102.frx":2106
      MousePointer    =   99  'Custom
      TabIndex        =   162
      ToolTipText     =   "Sila pastikan semua data telah dimasukkan dengan betul sebelum masukkan data ini ke dalam senarai jualan."
      Top             =   11040
      Width           =   1935
   End
   Begin VB.CommandButton CMD6 
      Caption         =   "Batal"
      Height          =   360
      Left            =   18000
      MouseIcon       =   "Frm102.frx":2410
      MousePointer    =   99  'Custom
      TabIndex        =   161
      ToolTipText     =   "Sila pastikan semua data telah dimasukkan dengan betul sebelum masukkan data ini ke dalam senarai trade in."
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton CMD4 
      Caption         =   "Masukkan Dalam Senarai Trade In"
      Height          =   360
      Left            =   15360
      MouseIcon       =   "Frm102.frx":271A
      MousePointer    =   99  'Custom
      TabIndex        =   160
      ToolTipText     =   "Sila pastikan semua data telah dimasukkan dengan betul sebelum masukkan data ini ke dalam senarai trade in."
      Top             =   2040
      Width           =   3135
   End
   Begin VB.CommandButton CMD5 
      Caption         =   "Masukkan Senarai Trade In"
      Height          =   360
      Left            =   15000
      MouseIcon       =   "Frm102.frx":2A24
      MousePointer    =   99  'Custom
      TabIndex        =   159
      ToolTipText     =   "Sila pastikan semua data telah dimasukkan dengan betul sebelum masukkan data ini ke dalam senarai trade in."
      Top             =   2040
      Width           =   2895
   End
   Begin VB.CommandButton CMD3 
      Caption         =   "Batal Edit Data"
      Height          =   360
      Left            =   11640
      MouseIcon       =   "Frm102.frx":2D2E
      MousePointer    =   99  'Custom
      TabIndex        =   158
      ToolTipText     =   "Sila pastikan semua data telah dimasukkan dengan betul sebelum masukkan data ini ke dalam senarai jualan."
      Top             =   2640
      Width           =   3135
   End
   Begin VB.CommandButton CMD1 
      Caption         =   "Masukkan Dalam Senarai Jualan"
      Height          =   360
      Left            =   8400
      MouseIcon       =   "Frm102.frx":3038
      MousePointer    =   99  'Custom
      TabIndex        =   157
      ToolTipText     =   "Sila pastikan semua data telah dimasukkan dengan betul sebelum masukkan data ini ke dalam senarai jualan."
      Top             =   2640
      Width           =   3135
   End
   Begin VB.CommandButton CMD2 
      Caption         =   "Masukkan Dalam Senarai Jualan"
      Height          =   360
      Left            =   8400
      MouseIcon       =   "Frm102.frx":3342
      MousePointer    =   99  'Custom
      TabIndex        =   156
      ToolTipText     =   "Sila pastikan semua data telah dimasukkan dengan betul sebelum masukkan data ini ke dalam senarai jualan."
      Top             =   2640
      Width           =   3135
   End
   Begin VB.CommandButton CMD7 
      Caption         =   "Carian Data"
      Height          =   360
      Left            =   2970
      MouseIcon       =   "Frm102.frx":364C
      MousePointer    =   99  'Custom
      TabIndex        =   155
      ToolTipText     =   "Carian data berkenaan no siri produk ini."
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Timer Tmr2 
      Interval        =   100
      Left            =   600
      Top             =   0
   End
   Begin VB.TextBox TB24 
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
      Height          =   300
      Left            =   17625
      TabIndex        =   146
      Text            =   "TB24"
      Top             =   1275
      Width           =   1245
   End
   Begin VB.TextBox TB9 
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   8550
      Locked          =   -1  'True
      TabIndex        =   145
      Text            =   "TB9"
      Top             =   5865
      Width           =   1500
   End
   Begin VB.TextBox TB8 
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5200
      Locked          =   -1  'True
      TabIndex        =   144
      Text            =   "TB8"
      Top             =   5865
      Width           =   1500
   End
   Begin VB.TextBox TB7 
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   142
      Text            =   "TB7"
      Top             =   5865
      Width           =   1500
   End
   Begin VB.ComboBox CBB4 
      BackColor       =   &H8000000A&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "Frm102.frx":3956
      Left            =   9915
      List            =   "Frm102.frx":3958
      Style           =   2  'Dropdown List
      TabIndex        =   139
      Top             =   9315
      Width           =   3990
   End
   Begin VB.TextBox TB23 
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   14760
      Locked          =   -1  'True
      TabIndex        =   127
      Text            =   "TB23"
      Top             =   8535
      Width           =   1140
   End
   Begin VB.TextBox TB21 
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   14760
      Locked          =   -1  'True
      TabIndex        =   124
      Text            =   "TB21"
      Top             =   7680
      Width           =   1140
   End
   Begin VB.TextBox TB20 
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   14760
      Locked          =   -1  'True
      TabIndex        =   121
      Text            =   "TB20"
      Top             =   7395
      Width           =   1140
   End
   Begin VB.TextBox TB19 
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
      Height          =   300
      Left            =   14760
      TabIndex        =   118
      Text            =   "TB19"
      Top             =   6855
      Width           =   1140
   End
   Begin VB.TextBox TB18 
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   10800
      Locked          =   -1  'True
      TabIndex        =   115
      Text            =   "TB18"
      Top             =   8295
      Width           =   1140
   End
   Begin VB.TextBox TB17 
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   10800
      Locked          =   -1  'True
      TabIndex        =   112
      Text            =   "TB17"
      Top             =   7995
      Width           =   1140
   End
   Begin VB.TextBox TB16 
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
      Height          =   300
      Left            =   10800
      TabIndex        =   109
      Text            =   "TB16"
      Top             =   7485
      Width           =   1140
   End
   Begin VB.TextBox TB15 
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
      Height          =   300
      Left            =   10800
      TabIndex        =   106
      Text            =   "TB15"
      Top             =   7170
      Width           =   1140
   End
   Begin VB.TextBox TB14 
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
      Height          =   300
      Left            =   10800
      TabIndex        =   103
      Text            =   "TB14"
      Top             =   6855
      Width           =   1140
   End
   Begin VB.TextBox TB22 
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
      Height          =   300
      Left            =   14760
      TabIndex        =   100
      Text            =   "TB22"
      Top             =   8295
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox TB12 
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   87
      Text            =   "TB12"
      Top             =   9720
      Width           =   1260
   End
   Begin VB.TextBox TB13 
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   86
      Text            =   "TB13"
      Top             =   10020
      Width           =   1260
   End
   Begin VB.CheckBox CB7 
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
      Left            =   2565
      TabIndex        =   83
      Top             =   9375
      Width           =   200
   End
   Begin VB.CheckBox CB5 
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
      TabIndex        =   82
      Top             =   9135
      Width           =   200
   End
   Begin VB.CheckBox CB6 
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
      Left            =   2280
      TabIndex        =   81
      Top             =   9150
      Width           =   200
   End
   Begin VB.TextBox TB11 
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
      Height          =   300
      Left            =   2760
      TabIndex        =   70
      Text            =   "TB11"
      Top             =   7845
      Width           =   1140
   End
   Begin VB.ComboBox CBB2 
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
      Height          =   315
      ItemData        =   "Frm102.frx":395A
      Left            =   17625
      List            =   "Frm102.frx":395C
      Style           =   2  'Dropdown List
      TabIndex        =   65
      Top             =   960
      Width           =   1245
   End
   Begin VB.TextBox TB10 
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
      Height          =   300
      Left            =   17625
      TabIndex        =   57
      Text            =   "TB10"
      Top             =   660
      Width           =   1245
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
      Height          =   300
      Left            =   2400
      TabIndex        =   50
      Text            =   "TB2"
      Top             =   2640
      Width           =   1140
   End
   Begin VB.TextBox TB6 
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   10680
      Locked          =   -1  'True
      TabIndex        =   40
      Text            =   "TB6"
      Top             =   2220
      Width           =   1260
   End
   Begin VB.TextBox TB5 
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   10680
      Locked          =   -1  'True
      TabIndex        =   37
      Text            =   "TB5"
      Top             =   1920
      Width           =   1260
   End
   Begin VB.CheckBox CB3 
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
      Left            =   10440
      TabIndex        =   32
      Top             =   1335
      Width           =   200
   End
   Begin VB.CheckBox CB2 
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
      Left            =   8280
      TabIndex        =   31
      Top             =   1320
      Width           =   200
   End
   Begin VB.CheckBox CB4 
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
      Left            =   10725
      TabIndex        =   30
      Top             =   1560
      Width           =   200
   End
   Begin VB.TextBox TB4 
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
      Height          =   300
      Left            =   10680
      TabIndex        =   26
      Text            =   "TB4"
      Top             =   795
      Width           =   1260
   End
   Begin VB.ComboBox CBB1 
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
      Height          =   315
      ItemData        =   "Frm102.frx":395E
      Left            =   6825
      List            =   "Frm102.frx":3960
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   2460
      Width           =   1245
   End
   Begin VB.TextBox TB3 
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
      Height          =   300
      Left            =   5880
      TabIndex        =   14
      Text            =   "TB3"
      Top             =   2160
      Width           =   1260
   End
   Begin VB.TextBox TB1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   1680
      TabIndex        =   2
      Text            =   "TB1"
      Top             =   1440
      Width           =   1260
   End
   Begin VB.CheckBox CB1 
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
      TabIndex        =   0
      Top             =   735
      Width           =   200
   End
   Begin VB.Timer Tmr1 
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2595
      Left            =   120
      TabIndex        =   53
      ToolTipText     =   "Sila ""Double Click"" untuk menu seterusnya."
      Top             =   3240
      Width           =   14685
      _ExtentX        =   25903
      _ExtentY        =   4577
      _Version        =   393216
      Rows            =   1
      Cols            =   0
      FixedCols       =   0
      BackColor       =   16777215
      ForeColor       =   0
      BackColorFixed  =   12648384
      ForeColorFixed  =   0
      BackColorSel    =   16777215
      ForeColorSel    =   12582912
      BackColorBkg    =   16777215
      GridColor       =   0
      WordWrap        =   -1  'True
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Height          =   3075
      Left            =   15000
      TabIndex        =   68
      ToolTipText     =   "Sila ""Double Click"" untuk menu seterusnya."
      Top             =   2760
      Width           =   6165
      _ExtentX        =   10874
      _ExtentY        =   5424
      _Version        =   393216
      Rows            =   1
      Cols            =   0
      FixedCols       =   0
      BackColor       =   16777215
      ForeColor       =   0
      BackColorFixed  =   12648384
      BackColorSel    =   16777215
      ForeColorSel    =   12582912
      BackColorBkg    =   16777215
      GridColor       =   0
      WordWrap        =   -1  'True
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   9915
      TabIndex        =   138
      Top             =   9000
      Width           =   3990
      _ExtentX        =   7038
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarTitleBackColor=   16744576
      Format          =   416415744
      CurrentDate     =   41561
   End
   Begin VB.Label L48_Text 
      Caption         =   "L48_Text"
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
      Left            =   14040
      TabIndex        =   154
      Top             =   5880
      Width           =   645
   End
   Begin VB.Label L46_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "L46_Text"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   11160
      TabIndex        =   153
      Top             =   10120
      Width           =   5145
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Nama  :"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   10440
      TabIndex        =   152
      Top             =   10120
      Width           =   825
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Maklumat agen."
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   8400
      TabIndex        =   151
      Top             =   9720
      Width           =   5865
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Bilangan barang trade in :"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   15000
      TabIndex        =   150
      Top             =   5880
      Width           =   2505
   End
   Begin VB.Label L44_Text 
      Caption         =   "L44_Text"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   17280
      TabIndex        =   149
      Top             =   5880
      Width           =   1005
   End
   Begin VB.Label L43_Text 
      Caption         =   "L43_Text"
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
      Left            =   12000
      TabIndex        =   148
      Top             =   5880
      Width           =   645
   End
   Begin VB.Label L34_Text 
      Caption         =   "Jumlah ini adalah nilai yang perlu dibayar oleh pihak kedai kepada agen."
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
      Height          =   525
      Left            =   4080
      TabIndex        =   147
      Top             =   8040
      Width           =   3765
   End
   Begin VB.Shape Shape5 
      Height          =   855
      Left            =   12045
      Top             =   7200
      Width           =   3975
   End
   Begin VB.Label Label112 
      BackStyle       =   0  'Transparent
      Caption         =   $"Frm102.frx":3962
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
      TabIndex        =   143
      Top             =   5880
      Width           =   14625
   End
   Begin VB.Label Label110 
      BackStyle       =   0  'Transparent
      Caption         =   "Tarikh Jualan      :"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   8280
      TabIndex        =   141
      Top             =   9000
      Width           =   2025
   End
   Begin VB.Label Label109 
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Pekerja     :"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   8280
      TabIndex        =   140
      Top             =   9315
      Width           =   2025
   End
   Begin VB.Shape Shape3 
      Height          =   855
      Left            =   8040
      Top             =   7800
      Width           =   3975
   End
   Begin VB.Label Label105 
      BackStyle       =   0  'Transparent
      Caption         =   "3) Maklumat harga jualan"
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
      Left            =   120
      TabIndex        =   137
      Top             =   6360
      Width           =   5385
   End
   Begin VB.Label Label108 
      BackStyle       =   0  'Transparent
      Caption         =   "4) Maklumat cara pembayaran"
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
      Left            =   8160
      TabIndex        =   136
      Top             =   6480
      Width           =   5385
   End
   Begin VB.Label Label107 
      BackStyle       =   0  'Transparent
      Caption         =   "Simpanan Duit Di Kedai Sebanyak : RM"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   12120
      TabIndex        =   135
      Top             =   8025
      Visible         =   0   'False
      Width           =   3435
   End
   Begin VB.Label L28_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "L28_Text"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   15525
      TabIndex        =   134
      Top             =   8025
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.Label L27_Text 
      Alignment       =   2  'Center
      Caption         =   "L27_Text"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   13920
      TabIndex        =   132
      Top             =   7200
      Width           =   600
   End
   Begin VB.Label L26_Text 
      Alignment       =   2  'Center
      Caption         =   "L26_Text"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   9960
      TabIndex        =   130
      Top             =   7800
      Width           =   600
   End
   Begin VB.Label Label101 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "RM :"
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
      Left            =   14130
      TabIndex        =   128
      Top             =   8520
      Width           =   600
   End
   Begin VB.Label Label99 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "RM :"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   14130
      TabIndex        =   125
      Top             =   7680
      Width           =   600
   End
   Begin VB.Label Label97 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "RM :"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   14130
      TabIndex        =   122
      Top             =   7395
      Width           =   600
   End
   Begin VB.Label Label95 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "RM :"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   14130
      TabIndex        =   119
      Top             =   6840
      Width           =   600
   End
   Begin VB.Label Label93 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "RM :"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   10170
      TabIndex        =   116
      Top             =   8280
      Width           =   600
   End
   Begin VB.Label Label91 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "RM :"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   10170
      TabIndex        =   113
      Top             =   7995
      Width           =   600
   End
   Begin VB.Label Label89 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "RM :"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   10170
      TabIndex        =   110
      Top             =   7485
      Width           =   600
   End
   Begin VB.Label Label87 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "RM :"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   10170
      TabIndex        =   107
      Top             =   7170
      Width           =   600
   End
   Begin VB.Label Label84 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "RM :"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   10170
      TabIndex        =   104
      Top             =   6840
      Width           =   600
   End
   Begin VB.Label Label82 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "RM :"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   14130
      TabIndex        =   101
      Top             =   8280
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label Label81 
      BackStyle       =   0  'Transparent
      Caption         =   "** Jumlah bayaran bagi emas dan upah."
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
      TabIndex        =   99
      Top             =   11175
      Width           =   4785
   End
   Begin VB.Label Label80 
      BackStyle       =   0  'Transparent
      Caption         =   "** Jumlah bayaran bagi upah SAHAJA."
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
      TabIndex        =   98
      Top             =   10695
      Width           =   4785
   End
   Begin VB.Label L13_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "L13_Text"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   4440
      TabIndex        =   97
      Top             =   10320
      Width           =   2715
   End
   Begin VB.Label Label78 
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah Upah"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   120
      TabIndex        =   96
      Top             =   10320
      Width           =   3225
   End
   Begin VB.Label Label77 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "RM :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   3600
      TabIndex        =   95
      Top             =   10320
      Width           =   825
   End
   Begin VB.Label L14_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "L14_Text"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   4440
      TabIndex        =   94
      Top             =   10800
      Width           =   2715
   End
   Begin VB.Label Label75 
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah Bayaran"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   120
      TabIndex        =   93
      Top             =   10800
      Width           =   3225
   End
   Begin VB.Label Label74 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "RM :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   3600
      TabIndex        =   92
      Top             =   10800
      Width           =   825
   End
   Begin VB.Label Label72 
      BackStyle       =   0  'Transparent
      Caption         =   "Harga Emas Dengan GST"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   120
      TabIndex        =   90
      Top             =   10035
      Width           =   2265
   End
   Begin VB.Label Label71 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "RM :"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2130
      TabIndex        =   89
      Top             =   9720
      Width           =   600
   End
   Begin VB.Label Label70 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "RM :"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2130
      TabIndex        =   88
      Top             =   10035
      Width           =   600
   End
   Begin VB.Label Label69 
      BackStyle       =   0  'Transparent
      Caption         =   "Harga Termasuk GST"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2835
      TabIndex        =   85
      Top             =   9330
      Width           =   2385
   End
   Begin VB.Label Label68 
      BackStyle       =   0  'Transparent
      Caption         =   "Zero Rated ZR(L)           Standard Rated SR"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   405
      TabIndex        =   84
      Top             =   9120
      Width           =   5610
   End
   Begin VB.Label Label66 
      BackStyle       =   0  'Transparent
      Caption         =   "Maklumat GST"
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
      Left            =   120
      TabIndex        =   80
      Top             =   8520
      Width           =   1575
   End
   Begin VB.Label Label65 
      BackStyle       =   0  'Transparent
      Caption         =   "** GST hanya dikenakan jika berat jualan emas kedai melebihi berat trade in."
      ForeColor       =   &H00000000&
      Height          =   525
      Left            =   1635
      TabIndex        =   79
      Top             =   8520
      Width           =   4905
   End
   Begin VB.Label Label63 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "RM :"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2130
      TabIndex        =   77
      Top             =   8160
      Width           =   600
   End
   Begin VB.Label L12_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "L12_Text"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2760
      TabIndex        =   76
      Top             =   8160
      Width           =   2145
   End
   Begin VB.Label L9_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "L9_Text"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   4440
      TabIndex        =   75
      Top             =   6600
      Width           =   3675
   End
   Begin VB.Label L10_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "L10_Text"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   4440
      TabIndex        =   74
      Top             =   6960
      Width           =   3675
   End
   Begin VB.Label L11_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "L11_Text"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   4440
      TabIndex        =   73
      Top             =   7320
      Width           =   3675
   End
   Begin VB.Label Label55 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "RM/g :"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2130
      TabIndex        =   71
      Top             =   7830
      Width           =   600
   End
   Begin VB.Label Label54 
      BackStyle       =   0  'Transparent
      Caption         =   "Senarai barang trade in."
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
      Left            =   15000
      TabIndex        =   69
      Top             =   2520
      Width           =   5385
   End
   Begin VB.Label Label56 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   17520
      TabIndex        =   66
      Top             =   960
      Width           =   150
   End
   Begin VB.Label L8_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "L8_Text"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   17625
      TabIndex        =   61
      Top             =   1605
      Width           =   1185
   End
   Begin VB.Label Label47 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   17520
      TabIndex        =   60
      Top             =   660
      Width           =   150
   End
   Begin VB.Label Label45 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   17520
      TabIndex        =   59
      Top             =   1260
      Width           =   150
   End
   Begin VB.Label Label44 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   17520
      TabIndex        =   58
      Top             =   1605
      Width           =   150
   End
   Begin VB.Label Label43 
      BackStyle       =   0  'Transparent
      Caption         =   "2) Maklumat barang trade in."
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
      Left            =   15000
      TabIndex        =   56
      Top             =   360
      Width           =   5385
   End
   Begin VB.Label Label42 
      BackStyle       =   0  'Transparent
      Caption         =   "1) Maklumat barang yang akan dijual."
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
      Left            =   120
      TabIndex        =   55
      Top             =   285
      Width           =   5385
   End
   Begin VB.Label Label41 
      BackStyle       =   0  'Transparent
      Caption         =   "Senarai barang yang dijual."
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
      TabIndex        =   54
      Top             =   3000
      Width           =   5385
   End
   Begin VB.Shape Shape2 
      Height          =   1095
      Left            =   120
      Top             =   1920
      Width           =   4095
   End
   Begin VB.Label Label39 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2265
      TabIndex        =   51
      Top             =   2670
      Width           =   150
   End
   Begin VB.Label Label38 
      BackStyle       =   0  'Transparent
      Caption         =   "Sila masukkan harga jualan semasa bagi emas dengan ketulenan 999.9 di bawah."
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
      Height          =   525
      Left            =   240
      TabIndex        =   49
      Top             =   2040
      Width           =   3945
   End
   Begin VB.Label Label37 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "(g) :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   3600
      TabIndex        =   48
      Top             =   7320
      Width           =   825
   End
   Begin VB.Label Label36 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "(g) :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   3600
      TabIndex        =   47
      Top             =   6960
      Width           =   825
   End
   Begin VB.Label Label35 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "(g) :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   3600
      TabIndex        =   46
      Top             =   6600
      Width           =   825
   End
   Begin VB.Label Label34 
      BackStyle       =   0  'Transparent
      Caption         =   "Beza Berat 999.9 "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   120
      TabIndex        =   45
      Top             =   7320
      Width           =   4305
   End
   Begin VB.Label Label33 
      BackStyle       =   0  'Transparent
      Caption         =   "Berat Belian 999.9"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   120
      TabIndex        =   44
      Top             =   6960
      Width           =   4185
   End
   Begin VB.Label Label30 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "RM :"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   10050
      TabIndex        =   41
      Top             =   2235
      Width           =   600
   End
   Begin VB.Label Label28 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "RM :"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   10050
      TabIndex        =   38
      Top             =   1920
      Width           =   600
   End
   Begin VB.Label Label27 
      BackStyle       =   0  'Transparent
      Caption         =   "** GST hanya dikenakan kepada upah SAHAJA."
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   9840
      TabIndex        =   36
      Top             =   1080
      Width           =   4185
   End
   Begin VB.Label Label50 
      BackStyle       =   0  'Transparent
      Caption         =   "Zero Rated ZR(L)           Standard Rated SR"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   8565
      TabIndex        =   35
      Top             =   1305
      Width           =   5610
   End
   Begin VB.Label Label49 
      BackStyle       =   0  'Transparent
      Caption         =   "Maklumat GST"
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
      Left            =   8325
      TabIndex        =   34
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label46 
      BackStyle       =   0  'Transparent
      Caption         =   "Harga Termasuk GST"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   10995
      TabIndex        =   33
      Top             =   1515
      Width           =   2385
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   6720
      TabIndex        =   29
      Top             =   2805
      Width           =   150
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   6720
      TabIndex        =   28
      Top             =   2505
      Width           =   150
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   10560
      TabIndex        =   27
      Top             =   795
      Width           =   150
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "Upah (RM)"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   8280
      TabIndex        =   25
      Top             =   795
      Width           =   1665
   End
   Begin VB.Label L6_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "L6_Text"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   5880
      TabIndex        =   24
      Top             =   1875
      Width           =   2745
   End
   Begin VB.Label L5_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "L5_Text"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   5880
      TabIndex        =   23
      Top             =   1560
      Width           =   3945
   End
   Begin VB.Label L4_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "L4_Text"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   5880
      TabIndex        =   22
      Top             =   1260
      Width           =   2745
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   5760
      TabIndex        =   21
      Top             =   2160
      Width           =   150
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   5760
      TabIndex        =   20
      Top             =   1875
      Width           =   150
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   5760
      TabIndex        =   19
      Top             =   1260
      Width           =   150
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   5760
      TabIndex        =   18
      Top             =   1560
      Width           =   150
   End
   Begin VB.Label L3_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "L3_Text"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   5880
      TabIndex        =   17
      Top             =   945
      Width           =   2745
   End
   Begin VB.Label L7_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "L7_Text"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   6825
      TabIndex        =   16
      Top             =   2805
      Width           =   3465
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   5760
      TabIndex        =   13
      Top             =   945
      Width           =   150
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Berat 999.9 (g)"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   4320
      TabIndex        =   12
      Top             =   2805
      Width           =   2505
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Kadar Tukaran Purity 999.9"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   4320
      TabIndex        =   11
      Top             =   2505
      Width           =   2505
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Berat Jualan (g)"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   4320
      TabIndex        =   10
      Top             =   2190
      Width           =   1665
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Berat Asal (g)"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   4320
      TabIndex        =   9
      Top             =   1875
      Width           =   1665
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Kategori Produk"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   4320
      TabIndex        =   8
      Top             =   1560
      Width           =   1665
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Purity"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   4320
      TabIndex        =   7
      Top             =   1260
      Width           =   1665
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Maklumat barang yang telah di scan."
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
      Left            =   4320
      TabIndex        =   6
      Top             =   600
      Width           =   5385
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "No. Siri Produk "
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   4320
      TabIndex        =   5
      Top             =   945
      Width           =   1665
   End
   Begin VB.Shape Shape1 
      Height          =   1335
      Left            =   120
      Top             =   600
      Width           =   4095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sila scan barang yang ingin dijual dalam ruangan di bawah."
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
      Height          =   525
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   3945
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "No. Siri Produk :"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   240
      TabIndex        =   3
      Top             =   1470
      Width           =   1545
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Scanner Mode"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   525
      TabIndex        =   1
      Top             =   705
      Width           =   1530
   End
   Begin VB.Label Label32 
      BackStyle       =   0  'Transparent
      Caption         =   "Berat Jualan 999.9"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   120
      TabIndex        =   43
      Top             =   6600
      Width           =   4185
   End
   Begin VB.Label Label31 
      BackStyle       =   0  'Transparent
      Caption         =   "Upah Dengan GST"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   8280
      TabIndex        =   42
      Top             =   2235
      Width           =   2265
   End
   Begin VB.Label Label29 
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah GST"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   8280
      TabIndex        =   39
      Top             =   1950
      Width           =   2505
   End
   Begin VB.Label Label40 
      BackStyle       =   0  'Transparent
      Caption         =   "Harga semasa (RM/g)"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   360
      TabIndex        =   52
      Top             =   2670
      Width           =   1905
   End
   Begin VB.Label Label52 
      BackStyle       =   0  'Transparent
      Caption         =   "Kadar Tukaran Purity 999.9"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   15120
      TabIndex        =   63
      Top             =   1260
      Width           =   2505
   End
   Begin VB.Label Label53 
      BackStyle       =   0  'Transparent
      Caption         =   "Berat (g)"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   15120
      TabIndex        =   64
      Top             =   660
      Width           =   1665
   End
   Begin VB.Label Label51 
      BackStyle       =   0  'Transparent
      Caption         =   "Berat 999.9 (g)"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   15120
      TabIndex        =   62
      Top             =   1605
      Width           =   2505
   End
   Begin VB.Label Label57 
      BackStyle       =   0  'Transparent
      Caption         =   "Purity Barang"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   15120
      TabIndex        =   67
      Top             =   960
      Width           =   2505
   End
   Begin VB.Label Label58 
      BackStyle       =   0  'Transparent
      Caption         =   "Harga semasa 999.9"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   120
      TabIndex        =   72
      Top             =   7830
      Width           =   2745
   End
   Begin VB.Label Label64 
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah Harga Emas"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   120
      TabIndex        =   78
      Top             =   8160
      Width           =   2505
   End
   Begin VB.Label Label73 
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah GST"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   120
      TabIndex        =   91
      Top             =   9750
      Width           =   2505
   End
   Begin VB.Label Label94 
      BackStyle       =   0  'Transparent
      Caption         =   "Potongan Kad Kredit"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   8160
      TabIndex        =   117
      Top             =   8280
      Width           =   2745
   End
   Begin VB.Label Label92 
      BackStyle       =   0  'Transparent
      Caption         =   "Cas Kad Kredit"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   8160
      TabIndex        =   114
      Top             =   7995
      Width           =   2745
   End
   Begin VB.Label Label90 
      BackStyle       =   0  'Transparent
      Caption         =   "Kad Kredit"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   8160
      TabIndex        =   111
      Top             =   7485
      Width           =   2745
   End
   Begin VB.Label Label88 
      BackStyle       =   0  'Transparent
      Caption         =   "Bank In"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   8160
      TabIndex        =   108
      Top             =   7170
      Width           =   2745
   End
   Begin VB.Label Label86 
      BackStyle       =   0  'Transparent
      Caption         =   "Tunai"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   8160
      TabIndex        =   105
      Top             =   6840
      Width           =   2745
   End
   Begin VB.Label Label100 
      BackStyle       =   0  'Transparent
      Caption         =   "Potongan Kad Debit"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   12120
      TabIndex        =   126
      Top             =   7680
      Width           =   2775
   End
   Begin VB.Label Label120 
      BackStyle       =   0  'Transparent
      Caption         =   "** Cas Kad Kredit :         %"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   8280
      TabIndex        =   131
      Top             =   7800
      Width           =   2760
   End
   Begin VB.Label Label96 
      BackStyle       =   0  'Transparent
      Caption         =   "Kad Debit"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   12120
      TabIndex        =   120
      Top             =   6840
      Width           =   2745
   End
   Begin VB.Label Label98 
      BackStyle       =   0  'Transparent
      Caption         =   "Cas Kad Debit"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   12120
      TabIndex        =   123
      Top             =   7395
      Width           =   2745
   End
   Begin VB.Label Label104 
      BackStyle       =   0  'Transparent
      Caption         =   "** Cas Kad Debit :         %"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   12240
      TabIndex        =   133
      Top             =   7200
      Width           =   2760
   End
   Begin VB.Label Label102 
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah Bayaran"
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
      Height          =   375
      Left            =   12120
      TabIndex        =   129
      Top             =   8520
      Width           =   2775
   End
   Begin VB.Label Label83 
      BackStyle       =   0  'Transparent
      Caption         =   "Duit Simpanan"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   12120
      TabIndex        =   102
      Top             =   8280
      Visible         =   0   'False
      Width           =   2745
   End
   Begin VB.Menu Frm102_PM_menu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu Frm102_SM_edit_data1 
         Caption         =   "Edit data"
      End
      Begin VB.Menu Frm102_SM_remove_jualan 
         Caption         =   "Keluarkan dari senarai jualan (Pulangkan ke stok kedai)"
      End
   End
   Begin VB.Menu Frm102_PM_menu2 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu Frm102_SM_edit_data2 
         Caption         =   "Edit data"
      End
      Begin VB.Menu Frm102_SM_remove_belian 
         Caption         =   "Keluarkan dari senarai belian"
      End
   End
End
Attribute VB_Name = "Frm102"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CB2_Click()
'On Error Resume Next
If Frm102.CB2 = 1 Then
    
    Frm102.CB3 = 0
    Frm102.CB4 = 0

End If

Call frm102_calc2
End Sub
Private Sub CB3_Click()
'On Error Resume Next
If Frm102.CB3 = 1 Then
    
    Frm102.CB2 = 0
    
End If

If Frm102.CB3 = 0 Then
    
    Frm102.CB4 = 0
    
End If

Call frm102_calc2
End Sub
Private Sub CB4_Click()
'On Error Resume Next
Call frm102_calc2
End Sub
Private Sub CB5_Click()
'On Error Resume Next
If Frm102.CB5 = 1 Then
    
    Frm102.CB6 = 0
    Frm102.CB7 = 0

End If

Call frm102_calc8
End Sub
Private Sub CB6_Click()
'On Error Resume Next
If Frm102.CB6 = 1 Then
    
    Frm102.CB5 = 0
    
End If

If Frm102.CB6 = 0 Then
    
    Frm102.CB7 = 0
    
End If

Call frm102_calc8
End Sub
Private Sub CB7_Click()
'On Error Resume Next
Call frm102_calc8
End Sub
Private Sub CBB1_Click()
'On Error Resume Next
Call frm102_calc1
End Sub
Private Sub CBB2_Click()
'On Error Resume Next
If Frm102.CBB2 <> vbNullString And GLOBAL_DISABLE = 0 Then
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from setting_database where Metal_Purity='" & Frm102.CBB2 & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
        If Not IsNull(rs!Kod_Metal_Purity) Then 'Kod Purity
            Frm102.L47_Text = rs!Kod_Metal_Purity
        Else
            Frm102.L47_Text = vbNullString
        End If
        If Not IsNull(rs!kadar_tukaran) Then
            If IsNumeric(rs!kadar_tukaran) Then
                Frm102.TB24 = rs!kadar_tukaran
            Else
                Frm102.TB24 = "0.00"
            End If
        Else
            Frm102.TB24 = "0.00"
        End If
    Else
        Frm102.TB24 = "0.00"
    End If
    
    rs.Close
    Set rs = Nothing
End If
End Sub
Private Sub CDM13_Click()
'On Error Resume Next
Frm102.Pic1.Visible = True
End Sub
Private Sub CMD1_Click()
'On Error Resume Next
Dim Err(30)
Dim Frm102_LM_BERAT_ASAL As Double
Dim Frm102_LM_BERAT_JUAL As Double
Dim Frm102_LM_HARGA_MODAL As Double
Dim Frm102_LM_HARGA_JUAL As Double
Dim Frm102_LM_HARGA_SEMASA_MODAL As Double
Dim Frm102_LM_TETAPANHARGA As Double
Dim Frm102_LM_LIMIT As Double
Dim Frm102_LM_HARGA_STAFF As Double 'Tetapan harga jualan kepada staff
Dim Frm102_LM_HARGA_PELANGGAN As Double 'Tetapan harga jualan kepada pelanggan
Dim Frm102_LM_HARGA_SEMASA As Double 'Harga semasa (jualan)
Dim Frm102_LM_BERAT_JUAL_ASAL As Double 'Berat Jualan (Purity Asal)
Dim Frm102_LM_HARGA_SEMASA_999 As Double 'Harga semasa (jualan) (Purity 999.9)
Dim Frm102_LM_HARGA_SUPPLIER As Double 'Harga per gram (harga semasa) dari supplier (modal)
Dim Frm102_LM_BERAT_999 As Double 'Berat Jualan (Purity Asal)
Dim Frm102_UPAH_MODAL As Double 'Upah modal
Dim Frm102_UPAH_JUAL As Double 'Upah jualan

Frm102_UPAH_MODAL = 0 'Upah modal
Frm102_UPAH_JUAL = 0 'Upah jualan
Frm102_LM_BERAT_JUAL_ASAL = 0 'Berat Jualan (Purity Asal)
Frm102_LM_HARGA_SEMASA_999 = 0 'Harga semasa (jualan) (Purity 999.9)
Frm102_LM_HARGA_SUPPLIER = 0 'Harga per gram (harga semasa) dari supplier (modal)
Frm102_LM_BERAT_999 = 0 'Berat Jualan (Purity Asal)

Frm102_LM_HARGA_SEMASA = 0 'Harga semasa (jualan)
Frm102_LM_HARGA_SUPPLIER = 0 'Harga per gram (harga semasa) dari supplier (modal)

x = 0
Frm102_LM_BERAT_ASAL = 0
Frm102_LM_BERAT_JUAL = 0
Frm102_LM_DATA_SAVE = 0
Frm102_LM_HARGA_MODAL = 0
Frm102_LM_HARGA_JUAL = 0
Frm102_LM_HARGA_SEMASA_MODAL = 0
Frm102_LM_PRICE_CHECK = 0 '0 : Tidak Perlu Periksa Harga Semasa Jualan , 1 : Perlu Periksa Harga Semasa Jualan
Frm102_LM_TETAPANHARGA = 0
Frm102_LM_LIMIT = 0
Frm102_LM_HARGA_STAFF = 0
Frm102_LM_HARGA_PELANGGAN = 0

If Frm102.L3_Text = vbNullString Then
    x = x + 1
    Err(x) = "Sila masukkan [No. Siri Produk]."
End If
If Frm102.L33_Text = vbNullString Then
    x = x + 1
    Err(x) = "Maklumat harga semasa modal belian item ini yang tidak sah. Sila keluar dari menu ini dan cuba sekali lagi."
End If
If Frm102.L50_Text = vbNullString Or (Frm102.L50_Text <> vbNullString And Not IsNumeric(Frm102.L50_Text)) Then
    x = x + 1
    Err(x) = "Maklumat upah modal yang tidak sah. Sila keluar dari menu ini dan cuba sekali lagi."
End If
If Frm102.L6_Text = vbNullString Or (Frm102.L6_Text <> vbNullString And Not IsNumeric(Frm102.L6_Text)) Then
    x = x + 1
    Err(x) = "Sila maklumat [Berat Asal]. Sila scan item sekali lagi."
End If
If Frm102.TB3 = vbNullString Or (Frm102.TB3 <> vbNullString And Not IsNumeric(Frm102.TB3)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Berat Jualan]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm102.TB2 = vbNullString Or (Frm102.TB2 <> vbNullString And Not IsNumeric(Frm102.TB2)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Harga Semasa Emas 999.9]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm102.TB2 <> vbNullString And IsNumeric(Frm102.TB2) Then

    If Format(Frm102.TB2, "0.00") = "0.00" Then
        x = x + 1
        Err(x) = "Harga emas semasa 999.9 yang tidak sah. Nilai 0.00 tidak dibenarkan."
    End If
    
End If
If Frm102.CBB1 = vbNullString Then
    x = x + 1
    Err(x) = "Sila pilih [Kadar Tukaran Purity 999.9]."
End If
If Frm102.L7_Text = vbNullString Or (Frm102.L7_Text <> vbNullString And Not IsNumeric(Frm102.L7_Text)) Then
    x = x + 1
    Err(x) = "[Berat 999.9] yang tidak sah. Sila scan item sekali lagi."
End If
If Frm102.TB4 = vbNullString Or (Frm102.TB4 <> vbNullString And Not IsNumeric(Frm102.TB4)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Upah]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm102.CB2 = 0 And Frm102.CB3 = 0 Then
    x = x + 1
    Err(x) = "Sila buat pilihan jenis GST bagi upah."
End If
If Frm102.TB5 = vbNullString Or Frm102.TB6 = vbNullString Then
    x = x + 1
    Err(x) = "Maklumat berkenaan GST yang tidak sah. Sila keluar dari menu ini dan cuba sekali lagi."
End If

If (Frm102.L6_Text <> vbNullString And IsNumeric(Frm102.L6_Text)) And (Frm102.TB3 <> vbNullString And IsNumeric(Frm102.TB3)) Then
    Frm102_LM_BERAT_ASAL = Frm102.L6_Text 'Berat Asal
    Frm102_LM_BERAT_JUAL = Frm102.TB3 'Berat Jualan
    
    If Frm102_LM_BERAT_JUAL > Frm102_LM_BERAT_ASAL Then
        x = x + 1
        Err(x) = "Berat jualan melebihi berat asal."
    End If
End If
If Frm102.L49_Text = vbNullString Or (Frm102.L49_Text <> vbNullString And Not IsNumeric(Frm102.L49_Text)) Then
    x = x + 1
    Err(x) = "Tiada maklumat harga dari supplier bagi purity ini." & vbCrLf & _
                "Sila pastikan harga dari supplier bagi purity ini telah ditetapkan dalam TETAPAN HARIAN SISTEM."
End If

If x <> 0 Then
    Frm6.Show
    Frm6.Pic1.Cls
    For Y = 1 To x
        Frm6.Pic1.Print Y & " - " & Err(Y)
    Next Y
    Exit Sub
Else
    Note = "Adakah anda ingin masukkan item ini ke dalam senarai jualan ?"
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbNo Then
        Exit Sub
    End If
    If Answer = vbYes Then
    
'### Periksa Data Dulang ### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from Data_Database where no_siri_Produk='" & Frm102.L3_Text & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Not IsNull(rs!dulang) Then Frm102_LM_DULANG = rs!dulang 'Dulang
        End If
        
        rs.Close
        Set rs = Nothing
'### Periksa Data Dulang ### - End
        
'### Masukkan Data Ke Dalam Temp Table ### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from " & G_JUALAN_TEMP & " where no_siri_Produk='" & Frm102.L3_Text & "'", cn, adOpenKeyset, adLockOptimistic
        
        If rs.EOF Then
            rs.AddNew
            If Frm102.L3_Text <> vbNullString Then
                rs!no_siri_Produk = Frm102.L3_Text 'No. Siri Produk
            Else
                rs!no_siri_Produk = Null 'No. Siri Produk
            End If
            If Frm102.L5_Text <> vbNullString Then
                rs!kategori_Produk = Frm102.L5_Text 'Kategori Produk
            Else
                rs!kategori_Produk = Null 'Kategori Produk
            End If
            If Frm102.L4_Text <> vbNullString Then
                rs!purity = Frm102.L4_Text 'Purity
            Else
                rs!purity = Null 'Purity
            End If
            If Frm102.L6_Text <> vbNullString Then
                rs!Berat_Asal = Format(Frm102.L6_Text, "0.00") 'Berat Asal (g)
            Else
                rs!Berat_Asal = Null 'Berat Asal (g)
            End If
            If Frm102.TB3 <> vbNullString Then
                rs!berat_jualan = Format(Frm102.TB3, "0.00") 'Berat Jualan (g)
            Else
                rs!berat_jualan = Null 'Berat Jualan (g)
            End If
            If Frm102.TB2 <> vbNullString Then
                rs!harga_Semasa = Format(Frm102.TB2, "0.00") 'Harga Semasa (RM/g)
            Else
                rs!harga_Semasa = Null 'Harga Semasa (RM/g)
            End If
            If Frm102.TB4 <> vbNullString Then
                rs!UPAH = Format(Frm102.TB4, "0.00") 'Upah (RM)
            Else
                rs!UPAH = Null 'Upah (RM)
            End If
            
            Frm102_LM_HARGA_SEMASA = Frm102.TB2 'Harga emas semasa 999.9 (Untuk tujuan jualan kepada pelanggan)
            Frm102_LM_BERAT_JUALAN_9999 = Frm102.L7_Text 'Berat jualan dalam purity 999.9
            Frm102_LM_UPAH_DAN_GST = Frm102.TB6 'Jumlah Upah + GST (Bagi jualan setiap item)

            If Frm102.TB6 <> vbNullString Then
                rs!harga_asal = Format((Frm102_LM_HARGA_SEMASA * Frm102_LM_BERAT_JUALAN_9999) + Frm102_LM_UPAH_DAN_GST, "0.00") 'Harga Asal Item (RM)
            Else
                rs!harga_asal = Null 'Harga Asal Item (RM)
            End If
            
            rs!diskaun = "0.00" 'Diskaun (%)
            rs!harga_lepas_diskaun = Format((Frm102_LM_HARGA_SEMASA * Frm102_LM_BERAT_JUALAN_9999) + Frm102_LM_UPAH_DAN_GST, "0.00") 'Harga Asal Item (RM)
            rs!adjustment = Format(0, "0.00") 'Adjustment (RM)
            rs!harga_jualan = Format((Frm102_LM_HARGA_SEMASA * Frm102_LM_BERAT_JUALAN_9999) + Frm102_LM_UPAH_DAN_GST, "0.00") 'Harga Asal Item (RM)
            rs!harga_jualan_dengan_gst = Format((Frm102_LM_HARGA_SEMASA * Frm102_LM_BERAT_JUALAN_9999) + Frm102_LM_UPAH_DAN_GST, "0.00") 'Harga Asal Item (RM)
            
            If Frm102.CB2 = 1 Then
                rs!gst_ari_nashi = "ZR (L)" '0 : Tiada GST : ZR(L) , 1 : Ada GST : SR
                rs!kadar_gst = Null 'Kadar Cukai GST (%)
                If Frm102.TB5 <> vbNullString Then
                    rs!jumlah_gst = Format(Frm102.TB5, "0.00") 'Jumlah GST (Bagi jualan setiap item)
                Else
                    rs!jumlah_gst = "0.00" 'Jumlah GST (Bagi jualan setiap item)
                End If
            ElseIf Frm102.CB3 = 1 Then
                rs!gst_ari_nashi = "SR" '0 : Tiada GST : ZR(L) , 1 : Ada GST : SR
                If Frm102.L21_Text <> vbNullString Then
                    rs!kadar_gst = Frm102.L21_Text 'Kadar Cukai GST (%)
                Else
                    rs!kadar_gst = "0" 'Jumlah Cukai GST (RM)
                End If
                If Frm102.TB5 <> vbNullString Then
                    rs!jumlah_gst = Format(Frm102.TB5, "0.00") 'Jumlah GST (Bagi jualan setiap item)
                Else
                    rs!jumlah_gst = "0.00" 'Jumlah GST (Bagi jualan setiap item)
                End If
                If Frm102.CB4 = 1 Then 'Jenis Cukai GST SR
                    rs!gst_include = 1 '0 : GST Dibayar Oleh Pelanggan , 1 : GST Termasuk Dalam Harga Barang
                Else
                    rs!gst_include = 0 '0 : GST Dibayar Oleh Pelanggan , 1 : GST Termasuk Dalam Harga Barang
                End If
            End If
            If Frm102.L30_Text <> vbNullString Then
                rs!harga_tanpa_gst = Format(Frm102.L30_Text, "0.00") 'Harga Jualan Tanpa GST (RM)
            Else
                rs!harga_tanpa_gst = Null 'Harga Jualan Tanpa GST (RM)
            End If
            If Frm102.TB6 <> vbNullString Then
                rs!harga_dengan_gst = Format(Frm102.TB6, "0.00") 'Harga Jualan Termasuk GST (RM)
            Else
                rs!harga_dengan_gst = Null 'Harga Jualan Termasuk GST (RM)
            End If
            rs!dropship = 0 '0 : Jualan Bukan Oleh Agen Dropship , 1 : Jualan Oleh Agen Dropship
            rs!komisyen_per_gram = Null 'Komisyen Per Gram Dropship (RM/g)
            rs!jumlah_komisyen = Null 'Jumlah Komisyen Kepada Agen Dropship Bagi Item Ini
            If Frm102.L32_Text = "0" Then '0 : Menu Data Baru , 1 : Menu Edit Data
                rs!Status = 1
            ElseIf Frm102.L32_Text = "1" Then
                rs!Status = 4
            End If
            rs!Type = 0 '0 : BK , 1 : Barang Permata
            If Frm102.L33_Text <> vbNullString Then
                rs!harga_per_gram_modal = Format(Frm102.L33_Text, "0.00") 'Harga Per Gram Bagi Modal (RM/g)
                Frm102_LM_HARGA_SEMASA_MODAL = Frm102.L33_Text
            Else
                rs!harga_per_gram_modal = Format(0, "0.00") 'Harga Per Gram Bagi Modal (RM/g)
            End If
            rs!modal = Format(Frm102_LM_HARGA_SEMASA_MODAL * Frm102_LM_BERAT_JUALAN_9999, "0.00") 'Harga Modal (RM)
            If IsNumeric(Frm102.TB6) And IsNumeric(Frm102.L33_Text) And IsNumeric(Frm102.TB3) Then
                Frm102_LM_HARGA_MODAL = Frm102.L33_Text * Frm102.TB3 'Harga modal
                Frm102_LM_HARGA_JUAL = (Frm102_LM_HARGA_SEMASA * Frm102_LM_BERAT_JUALAN_9999) + Frm102_LM_UPAH_DAN_GST 'Harga jualan
                
                rs!untung = Format(Frm102_LM_HARGA_JUAL - Frm102_LM_HARGA_MODAL, "0.00") 'Jumlah Keuntungan
            Else
                rs!untung = Format(0, "0.00") 'Jumlah Keuntungan
            End If
            
            If Frm102.L49_Text <> vbNullString Then 'Harga per gram (harga semasa) dari supplier (modal)
                rs!harga_per_gram_supplier = Frm102.L49_Text
            Else
                rs!harga_per_gram_supplier = 0
            End If
            
            If IsNumeric(Frm102.TB3) And IsNumeric(Frm102.TB2) And IsNumeric(Frm102.L49_Text) And IsNumeric(Frm102.L7_Text) And IsNumeric(Frm102.L6_Text) And IsNumeric(Frm102.L50_Text) And IsNumeric(Frm102.TB4) Then
                Frm102_LM_BERAT_JUAL_ASAL = Frm102.TB3 'Berat Jualan (Purity Asal)
                Frm102_LM_BERAT_ASAL = Frm102.L6_Text 'Berat Asal (Purity Asal)
                Frm102_UPAH_JUAL = Frm102.TB4 'Upah jualan
                Frm102_UPAH_MODAL = Frm102.L50_Text 'Upah modal
                Frm102_LM_HARGA_SEMASA_999 = Frm102.TB2 'Harga semasa (jualan) (Purity 999.9)
                Frm102_LM_HARGA_SUPPLIER = Frm102.L49_Text 'Harga per gram (harga semasa) dari supplier (modal)
                Frm102_LM_BERAT_999 = Frm102.L7_Text 'Berat emas dalam purity 999.9
                
                rs!upah_modal = Frm102.L50_Text 'Upah modal
                rs!harga_per_gram_supplier = Frm102.L49_Text 'Harga per gram (harga semasa) dari supplier (modal)
                rs!untung2 = Format(((Frm102_LM_BERAT_999 * Frm102_LM_HARGA_SEMASA_999) + Frm102_UPAH_JUAL) - ((Frm102_LM_BERAT_JUAL_ASAL * Frm102_LM_HARGA_SUPPLIER) + (Frm102_LM_BERAT_JUAL_ASAL / Frm102_LM_BERAT_ASAL) * Frm102_UPAH_MODAL), "0.00") 'Untung jika restok pada harga supplier ini
                
            Else
            
                rs!harga_per_gram_supplier = "0.00" 'Harga per gram (harga semasa) dari supplier (modal)
                rs!untung2 = "0.00" 'Untung jika restok pada harga supplier ini
                rs!upah_modal = "0.00" 'Upah modal
                
            End If
                
            If Format(Frm102.L6_Text, "0.00") = Format(Frm102.TB3, "0.00") Then
                rs!potong_flag = 0 '0 : Tiada Potong , 1 : Ada Potong
            Else
                rs!potong_flag = 1 '0 : Tiada Potong , 1 : Ada Potong
            End If
            rs!dulang = Frm102_LM_DULANG 'Dulang
            If Frm102.CBB1 <> vbNullString Then
                rs!pemalar_tukaran_999 = Frm102.CBB1 'Pemalar tukaran emas kepada 999.9 bagi urusan jualan
            Else
                rs!pemalar_tukaran_999 = Null 'Pemalar tukaran emas kepada 999.9 bagi urusan jualan
            End If
            If Frm102.L7_Text <> vbNullString Then
                rs!berat_999 = Format(Frm102.L7_Text, "0.00") 'Berat jualan dalam purity 999.9
            Else
                rs!berat_999 = Null 'Berat jualan dalam purity 999.9
            End If
            rs!gst_barang_atau_upah = 1 '0 : GST pada harga jualan , 1 : GST pada upah
            
            rs.Update
            Frm102_LM_DATA_SAVE = 1
        Else
            If Frm102.L3_Text <> vbNullString Then
                rs!no_siri_Produk = Frm102.L3_Text 'No. Siri Produk
            Else
                rs!no_siri_Produk = Null 'No. Siri Produk
            End If
            If Frm102.L5_Text <> vbNullString Then
                rs!kategori_Produk = Frm102.L5_Text 'Kategori Produk
            Else
                rs!kategori_Produk = Null 'Kategori Produk
            End If
            If Frm102.L4_Text <> vbNullString Then
                rs!purity = Frm102.L4_Text 'Purity
            Else
                rs!purity = Null 'Purity
            End If
            If Frm102.L6_Text <> vbNullString Then
                rs!Berat_Asal = Format(Frm102.L6_Text, "0.00") 'Berat Asal (g)
            Else
                rs!Berat_Asal = Null 'Berat Asal (g)
            End If
            If Frm102.TB3 <> vbNullString Then
                rs!berat_jualan = Format(Frm102.TB3, "0.00") 'Berat Jualan (g)
            Else
                rs!berat_jualan = Null 'Berat Jualan (g)
            End If
            If Frm102.TB2 <> vbNullString Then
                rs!harga_Semasa = Format(Frm102.TB2, "0.00") 'Harga Semasa (RM/g)
            Else
                rs!harga_Semasa = Null 'Harga Semasa (RM/g)
            End If
            If Frm102.TB4 <> vbNullString Then
                rs!UPAH = Format(Frm102.TB4, "0.00") 'Upah (RM)
            Else
                rs!UPAH = Null 'Upah (RM)
            End If
            
            Frm102_LM_HARGA_SEMASA = Frm102.TB2 'Harga emas semasa 999.9 (Untuk tujuan jualan kepada pelanggan)
            Frm102_LM_BERAT_JUALAN_9999 = Frm102.L7_Text 'Berat jualan dalam purity 999.9
            Frm102_LM_UPAH_DAN_GST = Frm102.TB6 'Jumlah Upah + GST (Bagi jualan setiap item)

            If Frm102.TB6 <> vbNullString Then
                rs!harga_asal = Format((Frm102_LM_HARGA_SEMASA * Frm102_LM_BERAT_JUALAN_9999) + Frm102_LM_UPAH_DAN_GST, "0.00") 'Harga Asal Item (RM)
            Else
                rs!harga_asal = Null 'Harga Asal Item (RM)
            End If
            
            rs!diskaun = "0.00" 'Diskaun (%)
            rs!harga_lepas_diskaun = Format((Frm102_LM_HARGA_SEMASA * Frm102_LM_BERAT_JUALAN_9999) + Frm102_LM_UPAH_DAN_GST, "0.00") 'Harga Asal Item (RM)
            rs!adjustment = Format(0, "0.00") 'Adjustment (RM)
            rs!harga_jualan = Format((Frm102_LM_HARGA_SEMASA * Frm102_LM_BERAT_JUALAN_9999) + Frm102_LM_UPAH_DAN_GST, "0.00") 'Harga Asal Item (RM)
            rs!harga_jualan_dengan_gst = Format((Frm102_LM_HARGA_SEMASA * Frm102_LM_BERAT_JUALAN_9999) + Frm102_LM_UPAH_DAN_GST, "0.00") 'Harga Asal Item (RM)
            
            If Frm102.CB2 = 1 Then
                rs!gst_ari_nashi = "ZR (L)" '0 : Tiada GST : ZR(L) , 1 : Ada GST : SR
                rs!kadar_gst = Null 'Kadar Cukai GST (%)
                If Frm102.TB5 <> vbNullString Then
                    rs!jumlah_gst = Format(Frm102.TB5, "0.00") 'Jumlah GST (Bagi jualan setiap item)
                Else
                    rs!jumlah_gst = "0.00" 'Jumlah GST (Bagi jualan setiap item)
                End If
            ElseIf Frm102.CB3 = 1 Then
                rs!gst_ari_nashi = "SR" '0 : Tiada GST : ZR(L) , 1 : Ada GST : SR
                If Frm102.L21_Text <> vbNullString Then
                    rs!kadar_gst = Frm102.L21_Text 'Kadar Cukai GST (%)
                Else
                    rs!kadar_gst = "0" 'Jumlah Cukai GST (RM)
                End If
                If Frm102.TB5 <> vbNullString Then
                    rs!jumlah_gst = Format(Frm102.TB5, "0.00") 'Jumlah GST (Bagi jualan setiap item)
                Else
                    rs!jumlah_gst = "0.00" 'Jumlah GST (Bagi jualan setiap item)
                End If
                If Frm102.CB4 = 1 Then 'Jenis Cukai GST SR
                    rs!gst_include = 1 '0 : GST Dibayar Oleh Pelanggan , 1 : GST Termasuk Dalam Harga Barang
                Else
                    rs!gst_include = 0 '0 : GST Dibayar Oleh Pelanggan , 1 : GST Termasuk Dalam Harga Barang
                End If
            End If
            If Frm102.L30_Text <> vbNullString Then
                rs!harga_tanpa_gst = Format(Frm102.L30_Text, "0.00") 'Harga Jualan Tanpa GST (RM)
            Else
                rs!harga_tanpa_gst = Null 'Harga Jualan Tanpa GST (RM)
            End If
            If Frm102.TB6 <> vbNullString Then
                rs!harga_dengan_gst = Format(Frm102.TB6, "0.00") 'Harga Jualan Termasuk GST (RM)
            Else
                rs!harga_dengan_gst = Null 'Harga Jualan Termasuk GST (RM)
            End If
            rs!dropship = 0 '0 : Jualan Bukan Oleh Agen Dropship , 1 : Jualan Oleh Agen Dropship
            rs!komisyen_per_gram = Null 'Komisyen Per Gram Dropship (RM/g)
            rs!jumlah_komisyen = Null 'Jumlah Komisyen Kepada Agen Dropship Bagi Item Ini
            If Frm102.L32_Text = "0" Then '0 : Menu Data Baru , 1 : Menu Edit Data
                rs!Status = 1
            ElseIf Frm102.L32_Text = "1" Then
                rs!Status = 3
            End If
            rs!Type = 0 '0 : BK , 1 : Barang Permata
            If Frm102.L33_Text <> vbNullString Then
                rs!harga_per_gram_modal = Format(Frm102.L33_Text, "0.00") 'Harga Per Gram Bagi Modal (RM/g)
                Frm102_LM_HARGA_SEMASA_MODAL = Frm102.L33_Text
            Else
                rs!harga_per_gram_modal = Format(0, "0.00") 'Harga Per Gram Bagi Modal (RM/g)
            End If
            rs!modal = Format(Frm102_LM_HARGA_SEMASA_MODAL * Frm102_LM_BERAT_JUALAN_9999, "0.00") 'Harga Modal (RM)
            If IsNumeric(Frm102.TB6) And IsNumeric(Frm102.L33_Text) And IsNumeric(Frm102.TB3) Then
                Frm102_LM_HARGA_MODAL = Frm102.L33_Text * Frm102.TB3 'Harga modal
                Frm102_LM_HARGA_JUAL = (Frm102_LM_HARGA_SEMASA * Frm102_LM_BERAT_JUALAN_9999) + Frm102_LM_UPAH_DAN_GST 'Harga jualan
                
                rs!untung = Format(Frm102_LM_HARGA_JUAL - Frm102_LM_HARGA_MODAL, "0.00") 'Jumlah Keuntungan
            Else
                rs!untung = Format(0, "0.00") 'Jumlah Keuntungan
            End If
            If Frm102.L49_Text <> vbNullString Then 'Harga per gram (harga semasa) dari supplier (modal)
                rs!harga_per_gram_supplier = Frm102.L49_Text
            Else
                rs!harga_per_gram_supplier = 0
            End If
            
            If IsNumeric(Frm102.TB3) And IsNumeric(Frm102.TB2) And IsNumeric(Frm102.L49_Text) And IsNumeric(Frm102.L7_Text) And IsNumeric(Frm102.L6_Text) And IsNumeric(Frm102.L50_Text) And IsNumeric(Frm102.TB4) Then
                Frm102_LM_BERAT_JUAL_ASAL = Frm102.TB3 'Berat Jualan (Purity Asal)
                Frm102_LM_BERAT_ASAL = Frm102.L6_Text 'Berat Asal (Purity Asal)
                Frm102_UPAH_JUAL = Frm102.TB4 'Upah jualan
                Frm102_UPAH_MODAL = Frm102.L50_Text 'Upah modal
                Frm102_LM_HARGA_SEMASA_999 = Frm102.TB2 'Harga semasa (jualan) (Purity 999.9)
                Frm102_LM_HARGA_SUPPLIER = Frm102.L49_Text 'Harga per gram (harga semasa) dari supplier (modal)
                Frm102_LM_BERAT_999 = Frm102.L7_Text 'Berat emas dalam purity 999.9
                
                rs!upah_modal = Frm102.L50_Text 'Upah modal
                rs!harga_per_gram_supplier = Frm102.L49_Text 'Harga per gram (harga semasa) dari supplier (modal)
                rs!untung2 = Format(((Frm102_LM_BERAT_999 * Frm102_LM_HARGA_SEMASA_999) + Frm102_UPAH_JUAL) - ((Frm102_LM_BERAT_JUAL_ASAL * Frm102_LM_HARGA_SUPPLIER) + (Frm102_LM_BERAT_JUAL_ASAL / Frm102_LM_BERAT_ASAL) * Frm102_UPAH_MODAL), "0.00") 'Untung jika restok pada harga supplier ini
                
            Else
            
                rs!harga_per_gram_supplier = "0.00" 'Harga per gram (harga semasa) dari supplier (modal)
                rs!untung2 = "0.00" 'Untung jika restok pada harga supplier ini
                rs!upah_modal = "0.00" 'Upah modal
                
            End If
            
            If Format(Frm102.L6_Text, "0.00") = Format(Frm102.TB3, "0.00") Then
                rs!potong_flag = 0 '0 : Tiada Potong , 1 : Ada Potong
            Else
                rs!potong_flag = 1 '0 : Tiada Potong , 1 : Ada Potong
            End If
            rs!dulang = Frm102_LM_DULANG 'Dulang
            If Frm102.CBB1 <> vbNullString Then
                rs!pemalar_tukaran_999 = Frm102.CBB1 'Pemalar tukaran emas kepada 999.9 bagi urusan jualan
            Else
                rs!pemalar_tukaran_999 = Null 'Pemalar tukaran emas kepada 999.9 bagi urusan jualan
            End If
            If Frm102.L7_Text <> vbNullString Then
                rs!berat_999 = Format(Frm102.L7_Text, "0.00") 'Berat jualan dalam purity 999.9
            Else
                rs!berat_999 = Null 'Berat jualan dalam purity 999.9
            End If
            rs!gst_barang_atau_upah = 1 '0 : GST pada harga jualan , 1 : GST pada upah
            
            rs.Update
            Frm102_LM_DATA_SAVE = 1
        End If
        
        rs.Close
        Set rs = Nothing
'### Masukkan Data Ke Dalam Temp Table ### - End
        
        If Frm102_LM_DATA_SAVE = 1 Then
            Call frm102_reset_1
            Call Frm102_Senarai_Jualan_Header
            Call Frm102_Senarai_Jualan
            
            MsgBox "Data telah berjaya dimasukkan ke dalam senarai jualan.", vbInformation, "Info"
            
            Frm102.TB1.SetFocus
        End If
    End If
End If
End Sub
Private Sub CMD10_Click()
'On Error Resume Next
Dim Err(30)
Dim rs1 As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim Frm102_LM_CUKAI_ZR As Double
Dim Frm102_LM_CUKAI_SR As Double
Dim Frm102_LM_BERAT_ASAL As Double
Dim Frm102_LM_BEZA_BERAT As Double
Dim Frm102_LM_BERAT_RETURN As Double
Dim Frm102_LM_BERAT_JUALAN As Double

Frm102_LM_KATEGORI = vbNullString
Frm102_LM_BERAT_ASAL = 0 'Berat Asal (g)
Frm102_LM_BERAT_JUALAN = 0 'Berat Jualan (g)
Frm102_LM_CUKAI_ZR = 0 'Jumlah cukai GST ZR
Frm102_LM_CUKAI_SR = 0 'Jumlah cukai GST SR
Frm102_LM_GENERATED = 0 '0 : Tiada No Voucher yang dihasilkan , 1 : Ada No. Voucher yang dihasilkan

Frm102_LM_Flag_SIMPANAN = 0 '0 : Tiada Penggunakan Duit Simpanan Kedai , 1 : Ada Penggunakan Duit Simpanan Kedai

If Frm102.TB14 = vbNullString Or (Frm102.TB14 <> vbNullString And Not IsNumeric(Frm102.TB14)) Then
    x = x + 1
    Err(x) = "Hanya NOMBOR dibenarkan dalam ruangan bayaran secara tunai. Sila masukkan 0 jika tiada bayaran tunai."
End If
If Frm102.TB15 = vbNullString Or (Frm102.TB15 <> vbNullString And Not IsNumeric(Frm102.TB15)) Then
    x = x + 1
    Err(x) = "Hanya NOMBOR dibenarkan dalam ruangan bayaran secara bank in. Sila masukkan 0 jika tiada bayaran bank in."
End If
If Frm102.TB16 = vbNullString Or (Frm102.TB16 <> vbNullString And Not IsNumeric(Frm102.TB16)) Then
    x = x + 1
    Err(x) = "Hanya NOMBOR dibenarkan dalam ruangan bayaran secara kad kredit. Sila masukkan 0 jika tiada bayaran kad kredit."
End If
If Frm102.TB22 = vbNullString Or (Frm102.TB22 <> vbNullString And Not IsNumeric(Frm102.TB22)) Then
    x = x + 1
    Err(x) = "Hanya NOMBOR dibenarkan dalam ruangan bayaran secara duit simpanan di kedai. Sila masukkan 0 jika tiada bayaran simpanan di kedai."
End If
If Frm102.TB19 = vbNullString Or (Frm102.TB19 <> vbNullString And Not IsNumeric(Frm102.TB19)) Then
    x = x + 1
    Err(x) = "Hanya NOMBOR dibenarkan dalam ruangan bayaran secara debit kad. Sila masukkan 0 jika tiada bayaran debit kad."
End If
If Format(Frm102.L14_Text, "0.00") <> Format(Frm102.TB23, "0.00") Then
    x = x + 1
    Err(x) = "Jumlah bayaran tidak sama dengan jumlah yang perlu dibayar."
End If
If Frm102.CBB4 = vbNullString Then
    x = x + 1
    Err(x) = "Sila pilih nama pekerja."
End If
If Frm102.L43_Text = 0 Then
    x = x + 1
    Err(x) = "Tiada senarai jualan."
End If
If Frm102.L44_Text = 0 Then
    x = x + 1
    Err(x) = "Tiada senarai barang trade in."
End If
'If Frm102.L46_Text = vbNullString Then
'    X = X + 1
'    Err(X) = "Tiada maklumat agen."
'End If
If Frm102.L34_Text.Visible = False Then
    If Frm102.TB11 = vbNullString Or (Frm102.TB11 <> vbNullString And Not IsNumeric(Frm102.TB11)) Then
        x = x + 1
        Err(x) = "Sila masukkan [Harga Semasa Emas 999.9]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm102.TB11 <> vbNullString And IsNumeric(Frm102.TB11) Then
    
        If Format(Frm102.TB11, "0.00") = "0.00" Then
            x = x + 1
            Err(x) = "Harga semasa emas 999.9 yang tidak sah. Nilai 0.00 tidak dibenarkan."
        End If
        
    End If

    If Frm102.CB5 = 0 And Frm102.CB6 = 0 Then
        x = x + 1
        Err(x) = "Sila buat pilihan cukai GST"
    End If
End If

If x <> 0 Then
    Frm6.Show
    Frm6.Pic1.Cls
    For Y = 1 To x
        Frm6.Pic1.Print Y & " - " & Err(Y)
    Next Y
    Exit Sub
Else

    If Frm102.L46_Text = vbNullString Then
    
        Note = "TIADA maklumat bagi agen yang diisi." & vbCrLf & _
                vbNullString & vbCrLf & _
                "Maklumat agen tidak akan dicetak di dalam invoice." & vbCrLf & _
                vbNullString & vbCrLf & _
                "Adakah anda yakin untuk teruskan urusan jualan ini ?"
    
    Else


        Note = "Adakah anda yakin untuk teruskan urusan jualan ini ?" & vbCrLf & _
                vbNullString & vbCrLf & _
                "Data jualan akan disimpan ke dalam sistem."

    End If
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbNo Then
        Exit Sub
    End If
    If Answer = vbYes Then
    
' ### Periksa kategori pembeli ### - Start
        If Frm102.L46_Text <> vbNullString Then
            If Frm28.L5_Text <> vbNullString Then

                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from senarai_pelanggan where no_pelanggan='" & Frm28.L5_Text & "'", cn, adOpenKeyset, adLockOptimistic

                If Not rs.EOF Then

                    If Not IsNull(rs!kategori_pelanggan) Then Frm102_LM_KATEGORI = rs!kategori_pelanggan

                End If

                rs.Close
                Set rs = Nothing

            End If
        End If
' ### Periksa kategori pembeli ### - End
        
        Frm102_LM_No_RESIT_JUALAN = Frm102.L23_Text 'No. Invoice
        Frm102_LM_No_VOUCHER = Frm102.L22_Text 'No. Voucher
        
    '###Masukkan Jualan Ke Dalam Table Akaun Jualan### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 22_jualan where no_resit='" & Frm102.L23_Text & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
        
            If Frm102.DTPicker1 <> vbNullString Then
                rs!tarikh = Frm102.DTPicker1 'Tarikh Jualan
            Else
                rs!tarikh = Null 'Tarikh Jualan
            End If
            
            rs!tunai = Format(0, "0.00")
            rs!bank_in = Format(0, "0.00") 'Cara Bayaran : Bank In
            rs!kad_kredit = Format(0, "0.00") 'Cara Bayaran : Kad Kredit
            rs!duit_simpanan_kedai = Format(0, "0.00") 'Cara Bayaran : Simpanan Duit Di Kedai
            rs!cek = Format(0, "0.00") 'Cara Bayaran : Cek
            
            '### Maklumat cara bayaran ### - Start
            If Frm102.TB14 <> vbNullString Then
                rs!tunai = Format(Frm102.TB14, "0.00") 'Cara Bayaran : Tunai
            Else
                rs!tunai = Null 'Cara Bayaran : Tunai
            End If
            If Frm102.TB15 <> vbNullString Then
                rs!bank_in = Format(Frm102.TB15, "0.00") 'Cara Bayaran : Bank In
            Else
                rs!bank_in = Null 'Cara Bayaran : Bank In
            End If
            If Frm102.TB16 <> vbNullString Then
                rs!kad_kredit = Format(Frm102.TB16, "0.00") 'Cara Bayaran : Kad Kredit
            Else
                rs!kad_kredit = Null 'Cara Bayaran : Kad Kredit
            End If
            If Frm102.L26_Text <> vbNullString Then
                rs!cas_Kad_Kredit = Frm102.L26_Text 'Cara Bayaran : Cas Kad Kredit (%)
            Else
                rs!cas_Kad_Kredit = Null 'Cara Bayaran : Cas Kad Kredit (%)
            End If
            If Frm102.TB17 <> vbNullString Then
                rs!jumlah_cas_kad_kredit = Format(Frm102.TB17, "0.00") 'Cara Bayaran : Jumlah Cas Kad Kredit (RM)
            Else
                rs!jumlah_cas_kad_kredit = Null 'Cara Bayaran : Jumlah Cas Kad Kredit (RM)
            End If
            If Frm102.TB18 <> vbNullString Then
                rs!jumlah_potongan_kad_kredit = Format(Frm102.TB18, "0.00") 'Cara Bayaran : Jumlah Potongan Kad Kredit (RM)
            Else
                rs!jumlah_potongan_kad_kredit = Null 'Cara Bayaran : Jumlah Potongan Kad Kredit (RM)
            End If
            If Frm102.TB22 <> vbNullString Then
                If Format(Frm102.TB22, "0.00") <> "0.00" Then
                    Frm102_LM_Flag_SIMPANAN = 1 '0 : Tiada Penggunakan Duit Simpanan Kedai , 1 : Ada Penggunakan Duit Simpanan Kedai
                End If
                rs!duit_simpanan_kedai = Format(Frm102.TB22, "0.00") 'Cara Bayaran : Simpanan Duit Di Kedai
            Else
                rs!duit_simpanan_kedai = Null 'Cara Bayaran : Simpanan Duit Di Kedai
            End If
            If Frm102.TB19 <> vbNullString Then
                rs!kad_debit = Format(Frm102.TB19, "0.00") 'Cara Bayaran : Kad Debit
            Else
                rs!kad_debit = Null 'Cara Bayaran : Kad Debit
            End If
            If Frm102.L27_Text <> vbNullString Then
                rs!cas_kad_debit = Frm102.L27_Text 'Cara Bayaran : Jumlah Cas Kad Debit (%)
            Else
                rs!cas_kad_debit = Null 'Cara Bayaran : Jumlah Cas Kad Debit (%)
            End If
            If Frm102.TB20 <> vbNullString Then
                rs!jumlah_cas_kad_debit = Format(Frm102.TB20, "0.00") 'Cara Bayaran : Jumlah Cas Kad Debit (RM)
            Else
                rs!jumlah_cas_kad_debit = Null 'Cara Bayaran : Jumlah Cas Kad Debit (RM)
            End If
            If Frm102.TB21 <> vbNullString Then
                rs!jumlah_potongan_kad_debit = Format(Frm102.TB21, "0.00") 'Cara Bayaran : Jumah Potongan Kad Debit (RM)
            Else
                rs!jumlah_potongan_kad_debit = Null 'Cara Bayaran : Jumah Potongan Kad Debit (RM)
            End If
            If Frm102.TB23 <> vbNullString Then
                rs!jumlah_bayaran = Format(Frm102.TB23, "0.00") 'Cara Bayaran : Jumlah Bayaran
            Else
                rs!jumlah_bayaran = Null 'Cara Bayaran : Jumlah Bayaran
            End If
            If Frm102.L15_Text <> vbNullString Then
                rs!harga_barang = Format(Frm102.L15_Text, "0.00") 'Jumlah Harga Barang Tanpa GST (RM)
            Else
                rs!harga_barang = Null 'Jumlah Harga Barang Tanpa GST (RM)
            End If
            If IsNumeric(Frm102.L19_Text) Then
                Frm102_LM_CUKAI_ZR = Frm102.L19_Text 'Maklumat GST : Jumlah GST ZR
            End If
            If IsNumeric(Frm102.L20_Text) Then
                Frm102_LM_CUKAI_SR = Frm102.L20_Text 'Maklumat GST : Jumlah GST SR
            End If
            rs!jumlah_cukai_gst = Format(Frm102_LM_CUKAI_ZR + Frm102_LM_CUKAI_SR, "0.00") 'Jumlah Cukai GST (ZR + SR)
            If Frm102.L16_Text <> vbNullString Then
                rs!harga_barang_dengan_gst = Format(Frm102.L16_Text, "0.00") 'Jumlah Harga Barang Dengan GST (RM)
                rs!harga_lepas_diskaun = Format(Frm102.L16_Text, "0.00") 'Harga Selepas Diskaun (RM)
                rs!harga_jualan = Format(Frm102.L16_Text, "0.00") 'Jumlah Harga Jualan (RM)
                rs!jumlah_perlu_bayar = Format(Frm102.L16_Text, "0.00") 'Jumlah Bayaran Yang Perlu Dibuat (RM)
            Else
                rs!harga_barang_dengan_gst = Null 'Jumlah Harga Barang Dengan GST (RM)
                rs!harga_lepas_diskaun = Null 'Harga Selepas Diskaun (RM)
                rs!harga_jualan = Null 'Jumlah Harga Jualan (RM)
                rs!jumlah_perlu_bayar = Null 'Jumlah Bayaran Yang Perlu Dibuat (RM)
            End If
            '### Maklumat cara bayaran ### - End
            
            rs!diskaun = Format(0, "0.00") 'Jumlah Diskaun (%)
            rs!adjustment = Format(0, "0.00") 'Adjustment (RM)
            rs!loss_trade_in = Format(0, "0.00") 'Potongan Harga Bagi Trade In Jika Kedai Perlu Bayar (%)
            rs!loss_trade_in_rm = Format(0, "0.00") 'Potongan Harga Bagi Trade In Jika Kedai Perlu Bayar (RM)
            rs!flag_bayaran = 0 '0 : Pembeli Bayar , 1 : Kedai Bayar
            rs!kuantiti_barang = Null 'Kuantiti Barang Yang Dijual
            rs!JUMLAH_BERAT = Null 'Jumlah Berat Barang Yang Dijual
            If Frm102.L17_Text <> vbNullString Then
                rs!gst_zr_harga = Format(Frm102.L17_Text, "0.00") 'Harga Keseluruhan Bagi Barang ZR
            Else
                rs!gst_zr_harga = Null 'Harga Keseluruhan Bagi Barang ZR
            End If
            If Frm102.L19_Text <> vbNullString Then
                rs!gst_zr_cukai = Format(Frm102.L19_Text, "0.00") 'Jumlah Cukai Bagi ZR
            Else
                rs!gst_zr_cukai = Null 'Jumlah Cukai Bagi ZR
            End If
            If Frm102.L18_Text <> vbNullString Then
                rs!gst_sr_harga = Format(Frm102.L18_Text, "0.00") 'Harga Keseluruhan Bagi Barang SR
            Else
                rs!gst_sr_harga = Null 'Harga Keseluruhan Bagi Barang SR
            End If
            If Frm102.L20_Text <> vbNullString Then
                rs!gst_sr_cukai = Format(Frm102.L20_Text, "0.00") 'Jumlah Cukai Bagi SR
            Else
                rs!gst_sr_cukai = Null 'Jumlah Cukai Bagi SR
            End If
            If Frm102.CBB4 <> vbNullString Then
                Frm102_LM_EMP_NO = Split(Frm102.CBB4, "  |  ")(1)
                rs!no_pekerja = Frm102_LM_EMP_NO 'No. Pekerja
            End If
            If Frm102.L46_Text <> vbNullString Then
                If Frm28.L5_Text <> vbNullString Then
                    rs!no_rujukan_pembeli = Frm28.L5_Text 'No. Rujukan Pembeli
                Else
                    rs!no_rujukan_pembeli = Null 'No. Rujukan Pembeli
                End If
            Else
                rs!no_rujukan_pembeli = Null 'No. Rujukan Pembeli
            End If
            If Frm102.L43_Text <> vbNullString Then
                rs!kuantiti_barang = Frm102.L43_Text 'Kuantiti barang
            Else
                rs!kuantiti_barang = 0 'Kuantiti barang
            End If
            If Frm102.L48_Text <> vbNullString Then
                rs!JUMLAH_BERAT = Frm102.L48_Text 'Jumlah berat
            Else
                rs!JUMLAH_BERAT = 0 'Kuantiti barang
            End If
            rs!no_rujukan_agen_dropship = Null 'No. Rujukan Agen Dropship
            rs!flag_trade_in = 0 '0 : Tiada Urusan Trade in , 1 : Ada Urusan Trade In
            rs!no_resit_trade_in = Null 'No. Resit Trade In
            rs!jumlah_trade_in = Null 'Jumlah Resit Trade In (RM)
            rs!jualan_online = 0
            
    '1:  Pelanggan
    '2:  Member
    '3:  RAF
    '4:  Pengedar
    '5:  Normal Dealer
    '6:  Master Dealer
    
            If Frm102_LM_KATEGORI <> vbNullString Then
                rs!kategori_pembeli = Frm102_LM_KATEGORI
            Else
                rs!kategori_pembeli = Null
            End If
            

            rs!invoice_type = 0 '0 : Unlimited , Selain 0 (Limited : Mengikut nombor yang dimasukkan)
            rs!epp = 0
            rs!approval_code_epp = Null
            rs!write_timestamp2 = Now
            DATA_SAVE = 1
            
            rs.Update

        End If
        
        rs.Close
        Set rs = Nothing
    '###Masukkan Jualan Ke Dalam Table Akaun Jualan### - End
    
'###Update Data Simpanan Duit Pelanggan### - Start
        'If Frm102_LM_Flag_SIMPANAN = 1 Then  '0 : Tiada Penggunakan Duit Simpanan Kedai , 1 : Ada Penggunakan Duit Simpanan Kedai
        
        '    Set rs = New ADODB.Recordset
        '    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        '    rs.Open "select * from senarai_pelanggan where no_pelanggan='" & Frm28.L5_Text & "'", cn, adOpenKeyset, adLockOptimistic
            
        '    If Not rs.EOF Then
        '        Frm102_LM_JUMLAH_SIMPANAN = Frm102.L26_Text  'Jumlah Simpanan Yang Ada
        '        Frm102_LM_GUNA_SIMPAN = Frm102.TB21  'Jumlah Simpanan Yang Hendak Digunakan
                
        '        rs!baki_simpanan = Format(Frm102_LM_JUMLAH_SIMPANAN - Frm102_LM_GUNA_SIMPAN, "0.00") 'Baki Simpanan
        '        rs.Update
        '    End If
            
        '    rs.Close
        '    Set rs = Nothing
            
        '    Set rs = New ADODB.Recordset
        '    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        '    rs.Open "select * from 24_rekod_kewangan_pelanggan", cn, adOpenKeyset, adLockOptimistic
            
        '    rs.AddNew
        '    rs!tarikh = Frm102.DTPicker1 'Tarikh
        '    rs!jenis = 1 '0 : Simpanan , 1 : Penggunaan Duit
        '    rs!no_rujukan_pelanggan = Frm28.L5_Text 'No. Rujukan Pelanggan
        '    rs!no_resit = "BK" & Format(Frm102_LM_No_RESIT_JUALAN, "000000") 'No. Resit Jualan
        '    rs!jumlah = Format(Frm102.TB22, "0.00") 'Jumlah Simpanan Yang Digunakan (RM)
        '    rs!jenis_penggunaan = 0 '0 : Belian Barangan Kemas , 1 : Ansuran , 2 : Tempahan (Deposit) , 3 : Servis , 4 : Tempahan (Ambilan Barang)
        '    rs.Update
            
        '    rs.Close
        '    Set rs = Nothing
           
        'End If
'###Update Data Simpanan Duit Pelanggan### - End

'### Masukkan Data Jualan Ke Dalam Table Jualan ### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from " & G_JUALAN_TEMP & "", cn, adOpenKeyset, adLockOptimistic
        
        While rs.EOF = False
        
            Frm102_LM_BERAT_ASAL = 0
            Frm102_LM_BEZA_BERAT = 0
            Frm102_LM_BERAT_RETURN = 0
            
'########### Kemasukan data baru dalam senarai ##############- Start
            If rs!Status = 4 Then

                Set rs1 = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs1.Open "select * from 23_senarai_jualan", cn, adOpenKeyset, adLockOptimistic
                
                rs1.AddNew
                rs1!tarikh = Frm102.DTPicker1 'Tarikh Jualan
                rs1!no_resit = Frm102.L23_Text 'No. Invoice Jualan
                If Not IsNull(rs!no_siri_Produk) Then
                    rs1!no_siri_Produk = rs!no_siri_Produk 'No. Siri Produk
                Else
                    rs1!no_siri_Produk = Null 'No. Siri Produk
                End If
                
                If Not IsNull(rs!kategori_Produk) Then
                    rs1!kategori_Produk = rs!kategori_Produk 'Kategori Produk
                Else
                    rs1!no_siri_Produk = Null 'Kategori Produk
                End If
                If Not IsNull(rs!purity) Then
                    rs1!purity = rs!purity 'Purity
                Else
                    rs1!purity = Null 'Purity
                End If
                If Not IsNull(rs!Berat_Asal) Then
                    rs1!Berat_Asal = rs!Berat_Asal 'Berat Asal (g)
                Else
                    rs1!Berat_Asal = Null 'Berat Asal (g)
                End If
                If Not IsNull(rs!berat_jualan) Then
                    rs1!berat_jualan = Format(rs!berat_jualan, "0.00") 'Berat Jualan (g)
                Else
                    rs1!berat_jualan = Null 'Berat Jualan (g)
                End If
                If Not IsNull(rs!harga_Semasa) Then
                    rs1!harga_Semasa = Format(rs!harga_Semasa, "0.00") 'Harga Semasa (RM/g)
                Else
                    rs1!harga_Semasa = Null 'Harga Semasa (RM/g)
                End If
                If Not IsNull(rs!UPAH) Then
                    rs1!UPAH = Format(rs!UPAH, "0.00") 'Upah (RM)
                Else
                    rs1!UPAH = Null 'Upah (RM)
                End If
                If Not IsNull(rs!harga_asal) Then
                    rs1!harga_asal = Format(rs!harga_asal, "0.00") 'Harga Asal Item (RM)
                Else
                    rs1!harga_asal = Null 'Harga Asal Item (RM)
                End If
                If Not IsNull(rs!diskaun) Then
                    rs1!diskaun = Format(rs!diskaun, "0.00") 'Diskaun (%)
                Else
                    rs1!diskaun = Null 'Diskaun (%)
                End If
                If Not IsNull(rs!harga_lepas_diskaun) Then
                    rs1!harga_lepas_diskaun = Format(rs!harga_lepas_diskaun, "0.00") 'Harga Selepas Diskaun (RM)
                Else
                    rs1!harga_lepas_diskaun = Null 'Harga Selepas Diskaun (RM)
                End If
                If Not IsNull(rs!adjustment) Then
                    rs1!adjustment = Format(rs!adjustment, "0.00") 'Harga Selepas Diskaun (RM)
                Else
                    rs1!adjustment = Null 'Harga Selepas Diskaun (RM)
                End If
                If Not IsNull(rs!harga_jualan) Then
                    rs1!harga_jualan = Format(rs!harga_jualan, "0.00") 'Harga Jualan (RM)
                Else
                    rs1!harga_jualan = Null 'Harga Jualan (RM)
                End If
                If Not IsNull(rs!gst_ari_nashi) Then
                    rs1!gst_ari_nashi = rs!gst_ari_nashi '0 : Tiada GST : ZR(L) , 1 : Ada GST : SR
                Else
                    rs1!gst_ari_nashi = Null '0 : Tiada GST : ZR(L) , 1 : Ada GST : SR
                End If
                If Not IsNull(rs!kadar_gst) Then
                    rs1!kadar_gst = rs!kadar_gst 'Kadar Cukai GST (%)
                Else
                    rs1!kadar_gst = Null 'Kadar Cukai GST (%)
                End If
                If Not IsNull(rs!jumlah_gst) Then
                    rs1!jumlah_gst = Format(rs!jumlah_gst, "0.00") 'Jumlah Cukai GST (RM)
                Else
                    rs1!jumlah_gst = Null 'Jumlah Cukai GST (RM)
                End If
                If Not IsNull(rs!harga_dengan_gst) Then
                    rs1!harga_dengan_gst = Format(rs!harga_dengan_gst, "0.00") 'Harga Jualan Termasuk GST (RM)
                Else
                    rs1!harga_dengan_gst = Null 'Harga Jualan Termasuk GST (RM)
                End If
                If Not IsNull(rs!dropship) Then
                    rs1!dropship = rs!dropship '0 : Jualan Bukan Oleh Agen Dropship , 1 : Jualan Oleh Agen Dropship
                Else
                    rs1!dropship = Null '0 : Jualan Bukan Oleh Agen Dropship , 1 : Jualan Oleh Agen Dropship
                End If
                If Not IsNull(rs!komisyen_per_gram) Then
                    rs1!komisyen_per_gram = Format(rs!komisyen_per_gram, "0.00") 'Komisyen Per Gram Dropship (RM/g) : BK
                Else
                    rs1!komisyen_per_gram = Null 'Komisyen Per Gram Dropship (RM/g) : BK
                End If
                If Not IsNull(rs!jumlah_komisyen) Then
                    rs1!jumlah_komisyen = Format(rs!jumlah_komisyen, "0.00") 'Jumlah Komisyen Kepada Agen Dropship Bagi Item Ini : BK
                Else
                    rs1!jumlah_komisyen = Null 'Jumlah Komisyen Kepada Agen Dropship Bagi Item Ini : BK
                End If
                If Not IsNull(rs!harga_per_gram_modal) Then
                    rs1!harga_per_gram_modal = Format(rs!harga_per_gram_modal, "0.00") 'Harga Per Gram Bagi Modal (RM/g)
                Else
                    rs1!harga_per_gram_modal = Null 'Harga Per Gram Bagi Modal (RM/g)
                End If
                If Not IsNull(rs!modal) Then
                    rs1!modal = Format(rs!modal, "0.00") 'Harga Modal (RM)
                Else
                    rs1!modal = Null 'Harga Modal (RM)
                End If
                If Not IsNull(rs!untung) Then
                    rs1!untung = Format(rs!untung, "0.00") 'Jumlah Keuntungan
                Else
                    rs1!untung = Null 'Jumlah Keuntungan
                End If
                If Not IsNull(rs!harga_per_gram_supplier) Then
                    rs1!harga_per_gram_supplier = Format(rs!harga_per_gram_supplier, "0.00") 'Harga per gram (harga semasa) dari supplier (modal)
                Else
                    rs1!harga_per_gram_supplier = Null 'Harga per gram (harga semasa) dari supplier (modal)
                End If
                If Not IsNull(rs!upah_modal) Then
                    rs1!upah_modal = Format(rs!upah_modal, "0.00") 'Upah modal
                Else
                    rs1!upah_modal = Null 'Upah modal
                End If
                If Not IsNull(rs!untung2) Then
                    rs1!untung2 = Format(rs!untung2, "0.00") 'Jumlah Keuntungan
                Else
                    rs1!untung2 = Null 'Jumlah Keuntungan
                End If
                If Not IsNull(rs!dulang) Then
                    rs1!dulang = rs!dulang 'Dulang
                Else
                    rs1!dulang = Null 'Dulang
                End If
                If Not IsNull(rs!potong_flag) Then
                    rs1!potong_flag = rs!potong_flag '0 : Tiada Potong , 1 : Ada Potong
                    If rs!potong_flag = 0 Then
                        rs1!Status = 0 '0 : Jualan Biasa , 1 : Jualan Secara Potong , 2 : Tempahan , 3 : Ansuran , 4 : ETA
                    Else
                        rs1!Status = 1 '0 : Jualan Biasa , 1 : Jualan Secara Potong , 2 : Tempahan , 3 : Ansuran , 4 : ETA
                    End If
                Else
                    rs1!potong_flag = Null '0 : Tiada Potong , 1 : Ada Potong
                End If

                If Not IsNull(rs!Type) Then
                    rs1!Type = rs!Type '0 : BK , 1 : Barang Permata
                Else
                    rs1!Type = Null '0 : BK , 1 : Barang Permata
                End
                rs1!jualan_online = 0
                If Frm102.CBB4 <> vbNullString Then
                    Frm102_LM_EMP_NO = Split(Frm102.CBB4, "  |  ")(1)
                    rs1!no_pekerja = Frm102_LM_EMP_NO 'No. Pekerja
                End If
                If Frm102.L46_Text <> vbNullString Then
                    If Frm28.L5_Text <> vbNullString Then
                        rs1!no_rujukan_pembeli = Frm28.L5_Text 'No. Rujukan Pembeli
                    Else
                        rs1!no_rujukan_pembeli = Null 'No. Rujukan Pembeli
                    End If
                Else
                    rs1!no_rujukan_pembeli = Null 'No. Rujukan Pembeli
                End If
                'If Frm27.L5_Text <> vbNullString Then
                '    rs1!no_rujukan_agen_dropship = Frm27.L5_Text 'No. Rujukan Agen Dropship
                'Else
                '    rs1!no_rujukan_agen_dropship = Null 'No. Rujukan Agen Dropship
                'End If
                
    '1:  Pelanggan
    '2:  Member
    '3:  RAF
    '4:  Pengedar
    '5:  Normal Dealer
    '6:  Master Dealer
    
                If Frm102_LM_KATEGORI <> vbNullString Then
                    rs1!kategori_pembeli = Frm102_LM_KATEGORI
                Else
                    rs1!kategori_pembeli = Null
                End If
                
                If Not IsNull(rs!gst_include) Then 'Pilihan Cukai GST (SR) Samada Pelanggan Bayar Atau Kedai Bayar
                    If rs!gst_include = 0 Then
                        rs1!gst_include = Null
                    ElseIf rs!gst_include = 1 Then
                        rs1!gst_include = "**Harga Termasuk GST"
                    End If
                Else
                    rs1!gst_include = Null
                End If
                If Not IsNull(rs!harga_tanpa_gst) Then
                    rs1!harga_tanpa_gst = Format(rs!harga_tanpa_gst, "0.00") 'Harga Semasa (RM/g)
                Else
                    rs1!harga_tanpa_gst = Null 'Harga Semasa (RM/g)
                End If

'### Maklumat tetapan harga jualan kepada staff ### - Start
                If Not IsNull(rs!kadar_penurunan_upah) Then 'Kadar peratusan penurunan harga upah kepada staff (%)
                    rs1!kadar_penurunan_upah = Format(rs!kadar_penurunan_upah, "0.00")
                Else
                    rs1!kadar_penurunan_upah = Null
                End If
                If Not IsNull(rs!harga_semasa_staff) Then 'Harga emas semasa yang dijual kepada staff
                    rs1!harga_semasa_staff = Format(rs!harga_semasa_staff, "0.00")
                Else
                    rs1!harga_semasa_staff = Null
                End If
                If Not IsNull(rs!kadar_penurunan_bp) Then 'Kadar peratusan penurunan harga barang permata kepada staff (%)
                    rs1!kadar_penurunan_bp = Format(rs!kadar_penurunan_bp, "0.00")
                Else
                    rs1!kadar_penurunan_bp = Null
                End If
                If Not IsNull(rs!harga_staff) Then 'Harga yang dijual kepada staff (RM)
                    rs1!harga_staff = Format(rs!harga_staff, "0.00")
                Else
                    rs1!harga_staff = Null
                End If
                If Not IsNull(rs!harga_bp_asal) Then 'Tetapan harga barang permata yang asal (RM)
                    rs1!harga_bp_asal = Format(rs!harga_bp_asal, "0.00")
                Else
                    rs1!harga_bp_asal = Null
                End If
                If Not IsNull(rs!upah_asal) Then 'Tetapan upah asal (RM)
                    rs1!upah_asal = Format(rs!upah_asal, "0.00")
                Else
                    rs1!upah_asal = Null
                End If
                If Not IsNull(rs!komisyen_staff) Then 'Tetapan upah asal (RM)
                    rs1!komisyen_staff = Format(rs!komisyen_staff, "0.00")
                Else
                    rs1!komisyen_staff = Null
                End If
'### Maklumat tetapan harga jualan kepada staff ### - End

                If Not IsNull(rs!pemalar_tukaran_999) Then 'Pemalar tukaran emas kepada 999.9 bagi urusan jualan
                    rs1!pemalar_tukaran_999 = rs!pemalar_tukaran_999
                Else
                    rs1!pemalar_tukaran_999 = Null
                End If
                If Not IsNull(rs!berat_999) Then 'Berat jualan dalam purity 999.9
                    rs1!berat_999 = Format(rs!berat_999, "0.00")
                Else
                    rs1!berat_999 = Null
                End If
                rs1!write_timestamp = Now
                rs1!jenis_jualan = 1 '0 : Jualan biasa kepada pelanggan , 1 : Jualan secara tukaran barang kepada agen
                If Not IsNull(rs!gst_barang_atau_upah) Then '0 : GST pada harga jualan , 1 : GST pada upah
                    rs1!gst_barang_atau_upah = rs!gst_barang_atau_upah
                Else
                    rs1!gst_barang_atau_upah = 0
                End If
                If Not IsNull(rs!harga_jualan_dengan_gst) Then
                    rs1!harga_jualan_dengan_gst = rs!harga_jualan_dengan_gst
                Else
                    rs1!harga_jualan_dengan_gst = 0
                End If
                
                rs1.Update
                
                rs1.Close
                Set rs1 = Nothing
            
'### Update Table Database Bagi Item Ini ### - Start
                Set rs2 = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs2.Open "select * from Data_Database where no_siri_produk='" & rs!no_siri_Produk & "'", cn, adOpenKeyset, adLockOptimistic
                
                If Not rs2.EOF Then
                    If rs!Type = 0 Then
                        Frm102_LM_BERAT_ASAL = rs2!beza_berat 'Berat Asal (g)
                        Frm102_LM_BERAT_JUALAN = rs!berat_jualan 'Berat Jualan (g)
                        
                        If Frm102_LM_BERAT_JUALAN = Frm102_LM_BERAT_ASAL Then
                            rs2!beza_berat = "0.00" 'Baki Berat
                            rs2!StatusItem = 11
                        Else
                            rs2!beza_berat = Format(Frm102_LM_BERAT_ASAL - Frm102_LM_BERAT_JUALAN, "0.00") 'Baki Berat
                            rs2!StatusItem = 12
                        End If
                    Else
                        rs2!StatusItem = 11
                    End If
                    rs2.Update
                End If

                rs2.Close
                Set rs2 = Nothing
'### Update Table Database Bagi Item Ini ### - End
            End If
'########### Kemasukan data baru dalam senarai ##############- End
            
'########### Edit data sedia ada dalam senarai ##############- Start
            ElseIf rs!Status = 3 Then
            
                Set rs1 = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs1.Open "select * from 23_senarai_jualan where ID='" & rs!id_database & "'", cn, adOpenKeyset, adLockOptimistic
                
                If Not rs1.EOF Then
                    rs1!tarikh = Frm102.DTPicker1 'Tarikh Jualan
                    rs1!no_resit = Frm102.L23_Text 'No. Invoice Jualan
                    If Not IsNull(rs!no_siri_Produk) Then
                        rs1!no_siri_Produk = rs!no_siri_Produk 'No. Siri Produk
                    Else
                        rs1!no_siri_Produk = Null 'No. Siri Produk
                    End If
                    If Not IsNull(rs!kategori_Produk) Then
                        rs1!kategori_Produk = rs!kategori_Produk 'Kategori Produk
                    Else
                        rs1!no_siri_Produk = Null 'Kategori Produk
                    End If
                    If Not IsNull(rs!purity) Then
                        rs1!purity = rs!purity 'Purity
                    Else
                        rs1!purity = Null 'Purity
                    End If
                    If Not IsNull(rs!Berat_Asal) Then
                        rs1!Berat_Asal = rs!Berat_Asal 'Berat Asal (g)
                    Else
                        rs1!Berat_Asal = Null 'Berat Asal (g)
                    End If
                    If Not IsNull(rs!berat_jualan) Then
                        rs1!berat_jualan = Format(rs!berat_jualan, "0.00") 'Berat Jualan (g)
                    Else
                        rs1!berat_jualan = Null 'Berat Jualan (g)
                    End If
                    If Not IsNull(rs!harga_Semasa) Then
                        rs1!harga_Semasa = Format(rs!harga_Semasa, "0.00") 'Harga Semasa (RM/g)
                    Else
                        rs1!harga_Semasa = Null 'Harga Semasa (RM/g)
                    End If
                    If Not IsNull(rs!UPAH) Then
                        rs1!UPAH = Format(rs!UPAH, "0.00") 'Upah (RM)
                    Else
                        rs1!UPAH = Null 'Upah (RM)
                    End If
                    If Not IsNull(rs!harga_asal) Then
                        rs1!harga_asal = Format(rs!harga_asal, "0.00") 'Harga Asal Item (RM)
                    Else
                        rs1!harga_asal = Null 'Harga Asal Item (RM)
                    End If
                    If Not IsNull(rs!diskaun) Then
                        rs1!diskaun = Format(rs!diskaun, "0.00") 'Diskaun (%)
                    Else
                        rs1!diskaun = Null 'Diskaun (%)
                    End If
                    If Not IsNull(rs!harga_lepas_diskaun) Then
                        rs1!harga_lepas_diskaun = Format(rs!harga_lepas_diskaun, "0.00") 'Harga Selepas Diskaun (RM)
                    Else
                        rs1!harga_lepas_diskaun = Null 'Harga Selepas Diskaun (RM)
                    End If
                    If Not IsNull(rs!adjustment) Then
                        rs1!adjustment = Format(rs!adjustment, "0.00") 'Harga Selepas Diskaun (RM)
                    Else
                        rs1!adjustment = Null 'Harga Selepas Diskaun (RM)
                    End If
                    If Not IsNull(rs!harga_jualan) Then
                        rs1!harga_jualan = Format(rs!harga_jualan, "0.00") 'Harga Jualan (RM)
                    Else
                        rs1!harga_jualan = Null 'Harga Jualan (RM)
                    End If
                    If Not IsNull(rs!gst_ari_nashi) Then
                        rs1!gst_ari_nashi = rs!gst_ari_nashi '0 : Tiada GST : ZR(L) , 1 : Ada GST : SR
                    Else
                        rs1!gst_ari_nashi = Null '0 : Tiada GST : ZR(L) , 1 : Ada GST : SR
                    End If
                    If Not IsNull(rs!kadar_gst) Then
                        rs1!kadar_gst = rs!kadar_gst 'Kadar Cukai GST (%)
                    Else
                        rs1!kadar_gst = Null 'Kadar Cukai GST (%)
                    End If
                    If Not IsNull(rs!jumlah_gst) Then
                        rs1!jumlah_gst = Format(rs!jumlah_gst, "0.00") 'Jumlah Cukai GST (RM)
                    Else
                        rs1!jumlah_gst = Null 'Jumlah Cukai GST (RM)
                    End If
                    If Not IsNull(rs!harga_dengan_gst) Then
                        rs1!harga_dengan_gst = Format(rs!harga_dengan_gst, "0.00") 'Harga Jualan Termasuk GST (RM)
                    Else
                        rs1!harga_dengan_gst = Null 'Harga Jualan Termasuk GST (RM)
                    End If
                    If Not IsNull(rs!dropship) Then
                        rs1!dropship = rs!dropship '0 : Jualan Bukan Oleh Agen Dropship , 1 : Jualan Oleh Agen Dropship
                    Else
                        rs1!dropship = Null '0 : Jualan Bukan Oleh Agen Dropship , 1 : Jualan Oleh Agen Dropship
                    End If
                    If Not IsNull(rs!komisyen_per_gram) Then
                        rs1!komisyen_per_gram = Format(rs!komisyen_per_gram, "0.00") 'Komisyen Per Gram Dropship (RM/g) : BK
                    Else
                        rs1!komisyen_per_gram = Null 'Komisyen Per Gram Dropship (RM/g) : BK
                    End If
                    If Not IsNull(rs!jumlah_komisyen) Then
                        rs1!jumlah_komisyen = Format(rs!jumlah_komisyen, "0.00") 'Jumlah Komisyen Kepada Agen Dropship Bagi Item Ini : BK
                    Else
                        rs1!jumlah_komisyen = Null 'Jumlah Komisyen Kepada Agen Dropship Bagi Item Ini : BK
                    End If
                    If Not IsNull(rs!harga_per_gram_modal) Then
                        rs1!harga_per_gram_modal = Format(rs!harga_per_gram_modal, "0.00") 'Harga Per Gram Bagi Modal (RM/g)
                    Else
                        rs1!harga_per_gram_modal = Null 'Harga Per Gram Bagi Modal (RM/g)
                    End If
                    If Not IsNull(rs!modal) Then
                        rs1!modal = Format(rs!modal, "0.00") 'Harga Modal (RM)
                    Else
                        rs1!modal = Null 'Harga Modal (RM)
                    End If
                    If Not IsNull(rs!untung) Then
                        rs1!untung = Format(rs!untung, "0.00") 'Jumlah Keuntungan
                    Else
                        rs1!untung = Null 'Jumlah Keuntungan
                    End If
                    If Not IsNull(rs!harga_per_gram_supplier) Then
                        rs1!harga_per_gram_supplier = Format(rs!harga_per_gram_supplier, "0.00") 'Harga per gram (harga semasa) dari supplier (modal)
                    Else
                        rs1!harga_per_gram_supplier = Null 'Harga per gram (harga semasa) dari supplier (modal)
                    End If
                    If Not IsNull(rs!upah_modal) Then
                        rs1!upah_modal = Format(rs!upah_modal, "0.00") 'Upah modal
                    Else
                        rs1!upah_modal = Null 'Upah modal
                    End If
                    If Not IsNull(rs!untung2) Then
                        rs1!untung2 = Format(rs!untung2, "0.00") 'Jumlah Keuntungan
                    Else
                        rs1!untung2 = Null 'Jumlah Keuntungan
                    End If
                    If Not IsNull(rs!dulang) Then
                        rs1!dulang = rs!dulang 'Dulang
                    Else
                        rs1!dulang = Null 'Dulang
                    End If
                    If Not IsNull(rs!potong_flag) Then
                        rs1!potong_flag = rs!potong_flag '0 : Tiada Potong , 1 : Ada Potong
                        If rs!potong_flag = 0 Then
                            rs1!Status = 0 '0 : Jualan Biasa , 1 : Jualan Secara Potong , 2 : Tempahan , 3 : Ansuran , 4 : ETA
                        Else
                            rs1!Status = 1 '0 : Jualan Biasa , 1 : Jualan Secara Potong , 2 : Tempahan , 3 : Ansuran , 4 : ETA
                        End If
                    Else
                        rs1!potong_flag = Null '0 : Tiada Potong , 1 : Ada Potong
                    End If
                    If Not IsNull(rs!Type) Then
                        rs1!Type = rs!Type '0 : BK , 1 : Barang Permata
                    Else
                        rs1!Type = Null '0 : BK , 1 : Barang Permata
                    End If
                    If Frm102.CBB4 <> vbNullString Then
                        Frm102_LM_EMP_NO = Split(Frm102.CBB4, "  |  ")(1)
                        rs1!no_pekerja = Frm102_LM_EMP_NO 'No. Pekerja
                    End If
                    rs1!jualan_online = 0
                    If Frm102.L46_Text <> vbNullString Then
                        If Frm28.L5_Text <> vbNullString Then
                            rs1!no_rujukan_pembeli = Frm28.L5_Text 'No. Rujukan Pembeli
                        Else
                            rs1!no_rujukan_pembeli = Null 'No. Rujukan Pembeli
                        End If
                    Else
                        rs1!no_rujukan_pembeli = Null 'No. Rujukan Pembeli
                    End If
        '1:  Pelanggan
        '2:  Member
        '3:  RAF
        '4:  Pengedar
        '5:  Normal Dealer
        '6:  Master Dealer
        
                    If Frm102_LM_KATEGORI <> vbNullString Then
                        rs1!kategori_pembeli = Frm102_LM_KATEGORI
                    Else
                        rs1!kategori_pembeli = Null
                    End If

                    If Not IsNull(rs!gst_include) Then 'Pilihan Cukai GST (SR) Samada Pelanggan Bayar Atau Kedai Bayar
                        If rs!gst_include = 0 Then
                            rs1!gst_include = Null
                        ElseIf rs!gst_include = 1 Then
                            rs1!gst_include = "**Harga Termasuk GST"
                        End If
                    Else
                        rs1!gst_include = Null
                    End If
                    If Not IsNull(rs!harga_tanpa_gst) Then
                        rs1!harga_tanpa_gst = Format(rs!harga_tanpa_gst, "0.00") 'Harga Semasa (RM/g)
                    Else
                        rs1!harga_tanpa_gst = Null 'Harga Semasa (RM/g)
                    End If
    
    '### Maklumat tetapan harga jualan kepada staff ### - Start
                    If Not IsNull(rs!kadar_penurunan_upah) Then 'Kadar peratusan penurunan harga upah kepada staff (%)
                        rs1!kadar_penurunan_upah = Format(rs!kadar_penurunan_upah, "0.00")
                    Else
                        rs1!kadar_penurunan_upah = Null
                    End If
                    If Not IsNull(rs!harga_semasa_staff) Then 'Harga emas semasa yang dijual kepada staff
                        rs1!harga_semasa_staff = Format(rs!harga_semasa_staff, "0.00")
                    Else
                        rs1!harga_semasa_staff = Null
                    End If
                    If Not IsNull(rs!kadar_penurunan_bp) Then 'Kadar peratusan penurunan harga barang permata kepada staff (%)
                        rs1!kadar_penurunan_bp = Format(rs!kadar_penurunan_bp, "0.00")
                    Else
                        rs1!kadar_penurunan_bp = Null
                    End If
                    If Not IsNull(rs!harga_staff) Then 'Harga yang dijual kepada staff (RM)
                        rs1!harga_staff = Format(rs!harga_staff, "0.00")
                    Else
                        rs1!harga_staff = Null
                    End If
                    If Not IsNull(rs!harga_bp_asal) Then 'Tetapan harga barang permata yang asal (RM)
                        rs1!harga_bp_asal = Format(rs!harga_bp_asal, "0.00")
                    Else
                        rs1!harga_bp_asal = Null
                    End If
                    If Not IsNull(rs!upah_asal) Then 'Tetapan upah asal (RM)
                        rs1!upah_asal = Format(rs!upah_asal, "0.00")
                    Else
                        rs1!upah_asal = Null
                    End If
                    If Not IsNull(rs!komisyen_staff) Then 'Tetapan upah asal (RM)
                        rs1!komisyen_staff = Format(rs!komisyen_staff, "0.00")
                    Else
                        rs1!komisyen_staff = Null
                    End If
    '### Maklumat tetapan harga jualan kepada staff ### - End
    
                    If Not IsNull(rs!pemalar_tukaran_999) Then 'Pemalar tukaran emas kepada 999.9 bagi urusan jualan
                        rs1!pemalar_tukaran_999 = rs!pemalar_tukaran_999
                    Else
                        rs1!pemalar_tukaran_999 = Null
                    End If
                    If Not IsNull(rs!berat_999) Then 'Berat jualan dalam purity 999.9
                        rs1!berat_999 = Format(rs!berat_999, "0.00")
                    Else
                        rs1!berat_999 = Null
                    End If
                    If Not IsNull(rs!gst_barang_atau_upah) Then '0 : GST pada harga jualan , 1 : GST pada upah
                        rs1!gst_barang_atau_upah = rs!gst_barang_atau_upah
                    Else
                        rs1!gst_barang_atau_upah = 0
                    End If
                    If Not IsNull(rs!harga_jualan_dengan_gst) Then
                        rs1!harga_jualan_dengan_gst = rs!harga_jualan_dengan_gst
                    Else
                        rs1!harga_jualan_dengan_gst = 0
                    End If
                    rs1!write_timestamp2 = Now
                    
                    rs1.Update
                End If

                rs1.Close
                Set rs1 = Nothing
                               
'### Update Table Database Bagi Item Ini ### - Start
                Set rs2 = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs2.Open "select * from Data_Database where no_siri_produk='" & rs!no_siri_Produk & "'", cn, adOpenKeyset, adLockOptimistic
                
                If Not rs2.EOF Then
                    If rs!Type = 0 Then
                        Frm102_LM_BERAT_ASAL = Format(rs2!Berat, "0.00") 'Berat Asal (g)
                        Frm102_LM_BEZA_BERAT = Format(rs2!beza_berat, "0.00") 'Berat Asal (g)
                        Frm102_BERAT_JUALAN_BARU = Format(rs!berat_jualan, "0.00") 'Berat Jualan (g)
                        
                        Frm102_LM_BAKI_BERAT = Frm102_BERAT_JUALAN_BARU - Format((Frm102_LM_BERAT_JUALAN_ASAL + Frm102_LM_BEZA_BERAT), "0.00")
                        
                        If Frm102_LM_BAKI_BERAT = 0 Then
                            rs2!beza_berat = "0.00" 'Baki Berat
                            rs2!StatusItem = 11
                        Else
                            rs2!beza_berat = Format(Frm102_LM_BERAT_ASAL - Frm102_BERAT_JUALAN_BARU, "0.00") 'Baki Berat
                            rs2!StatusItem = 12
                        End If
                    Else
                        rs2!StatusItem = 11
                    End If
                    rs2.Update
                End If

                rs2.Close
                Set rs2 = Nothing
'### Update Table Database Bagi Item Ini ### - End

'########### Edit data sedia ada dalam senarai ##############- End
            
            ElseIf rs!Status = 5 Then

                Set rs1 = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs1.Open "select * from 23_senarai_jualan where ID='" & rs!id_database & "'", cn, adOpenKeyset, adLockOptimistic
            
                If Not rs1.EOF Then
                    If Not IsNull(rs1!berat_jualan) Then
                        Frm102_LM_BERAT_RETURN = rs1!berat_jualan
                    End If
                    rs1.Delete
                    rs1.Update
                End If
                
                rs1.Close
                Set rs1 = Nothing
                
'### Update Table Database Bagi Item Ini ### - Start
                Set rs2 = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs2.Open "select * from Data_Database where no_siri_produk='" & rs!no_siri_Produk & "'", cn, adOpenKeyset, adLockOptimistic
                
                If Not rs2.EOF Then
                    If rs!Type = 0 Then
                        Frm102_LM_BERAT_ASAL = rs2!Berat 'Berat Asal (g)
                        Frm102_LM_BEZA_BERAT = rs2!beza_berat 'Berat Asal (g)
                        
                        If Frm102_LM_BERAT_RETURN + Frm102_LM_BEZA_BERAT = Frm102_LM_BERAT_ASAL Then
                            rs2!beza_berat = Format(Frm102_LM_BERAT_ASAL, "0.00") 'Baki Berat
                            rs2!StatusItem = 10
                        Else
                            rs2!beza_berat = Format(Frm102_LM_BERAT_RETURN + Frm102_LM_BEZA_BERAT, "0.00") 'Baki Berat
                            rs2!StatusItem = 12
                        End If
                    Else
                        rs2!StatusItem = 10
                    End If
                    rs2.Update
                End If
                
                rs2.Close
                Set rs2 = Nothing
'### Update Table Database Bagi Item Ini ### - End
            
            End If

            
            rs.MoveNext
        Wend '2
        
        rs.Close
        Set rs = Nothing

'### Masukkan data belian barang dari agen ke dalam database ### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 49_belian_temp", cn, adOpenKeyset, adLockOptimistic
        
        While rs.EOF = False
        
        '### Kemasukkan data baru ### - Start
            If rs!Status = 4 Then
            
                Set rs1 = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs1.Open "select * from 50_belian_emas_agen", cn, adOpenKeyset, adLockOptimistic
            
                rs1.AddNew
                rs1!no_invoice = Frm102.L23_Text 'No. Invoice
                rs1!tarikh = Frm102.DTPicker1 'Tarikh Jualan
                If Not IsNull(rs!Berat_Asal) Then rs1!Berat_Asal = rs!Berat_Asal 'Berat asal barang
                If Not IsNull(rs!purity) Then rs1!purity = rs!purity 'Purity barang
                If Not IsNull(rs!kadar_tukaran) Then rs1!kadar_tukaran = rs!kadar_tukaran 'Kadar tukaran kepada purity 999.9
                If Not IsNull(rs!berat_tukaran) Then rs1!berat_tukaran = rs!berat_tukaran 'Berat setelah ditukar kepada purity 999.9
                If Not IsNull(rs!Status) Then rs1!Status = 1 'Status
                rs1!write_timestamp = Now
                
                rs1.Update
                
                rs1.Close
                Set rs1 = Nothing
                
            ElseIf rs!Status = 3 Then
                
                Set rs1 = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs1.Open "select * from 50_belian_emas_agen where ID='" & rs!id_database & "'", cn, adOpenKeyset, adLockOptimistic
            
                If Not rs1.EOF Then
                    rs1!no_invoice = Frm102.L23_Text 'No. Invoice
                    rs1!tarikh = Frm102.DTPicker1 'Tarikh Jualan
                    If Not IsNull(rs!Berat_Asal) Then rs1!Berat_Asal = rs!Berat_Asal 'Berat asal barang
                    If Not IsNull(rs!purity) Then rs1!purity = rs!purity 'Purity barang
                    If Not IsNull(rs!kadar_tukaran) Then rs1!kadar_tukaran = rs!kadar_tukaran 'Kadar tukaran kepada purity 999.9
                    If Not IsNull(rs!berat_tukaran) Then rs1!berat_tukaran = rs!berat_tukaran 'Berat setelah ditukar kepada purity 999.9
                    If Not IsNull(rs!Status) Then rs1!Status = 1 'Status
                    rs1!write_timestamp2 = Now
                    
                    rs1.Update
                End If
                
                rs1.Close
                Set rs1 = Nothing
            
            ElseIf rs!Status = 5 Then
            
                Set rs1 = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs1.Open "select * from 50_belian_emas_agen where ID='" & rs!id_database & "'", cn, adOpenKeyset, adLockOptimistic
            
                If Not rs1.EOF Then
                    rs1!Status = 0 'Status
                    rs1!write_timestamp2 = Now
                    rs1.Update
                End If
                
                rs1.Close
                Set rs1 = Nothing
            
            End If
            
            rs.MoveNext
        Wend '1
        
        rs.Close
        Set rs = Nothing
'### Masukkan data belian barang dari agen ke dalam database ### - End
        
'### Periksa samada ada pembayaran kepada agen atau tidak bagi urusan belian barang kemas ### - Start
'Frm102.L45_Text = 0 'Flag bagi jika ada pengeluaran voucher bagi urusan ini , 0 : Tiada voucher / Tiada history pengeluaran voucher , 1 : Ada voucher / Ada history pengeluaran voucher
        If Frm102.L34_Text.Visible = True And Frm102.L45_Text = 0 Then
            
            Frm102_LM_GENERATED = 1 '0 : Tiada No Voucher yang dihasilkan , 1 : Ada No. Voucher yang dihasilkan
            
Re_Gen_No_Rujukan2:
'###Masukkan Jualan Ke Dalam Table Akaun Jualan### - Start
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from 51_voucher_belian_agen where no_voucher='" & "TIA" & Format(Frm102_LM_No_VOUCHER, "000000") & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                Frm102_LM_No_VOUCHER = Frm102_LM_No_VOUCHER + 1
                Frm102.L22_Text = Frm102_LM_No_VOUCHER 'No. Invoice Jualan
                
                rs.Close
                Set rs = Nothing
                
                GoTo Re_Gen_No_Rujukan2:
            End If
            
            rs.Close
            Set rs = Nothing
        End If
'### Periksa samada ada pembayaran kepada agen atau tidak bagi urusan belian barang kemas ### - End

'### Masukkan data voucher / invoice bagi belian agen ini ### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 51_voucher_belian_agen where no_invoice='" & Frm102.L23_Text & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Frm102.L34_Text.Visible = True Then 'No. Voucher
                If Frm102_LM_GENERATED = 1 Then rs!no_voucher = "TIA" & Format(Frm102_LM_No_VOUCHER, "000000")
                rs!flag_bayaran = 1 '0 : Bayaran dibuat oleh pembeli , 1 : Bayaran dibuat oleh pihak kedai
            Else
                'rs!no_voucher = Null
                rs!flag_bayaran = 0 '0 : Bayaran dibuat oleh pembeli , 1 : Bayaran dibuat oleh pihak kedai
            End If
            rs!tarikh = Frm102.DTPicker1 'Tarikh belian
            If Frm102.L9_Text <> vbNullString Then 'Berat jualan keseluruhan barang kedai
                rs!berat_jualan = Format(Frm102.L9_Text, "0.00")
            Else
                rs!berat_jualan = "0.00"
            End If
            If Frm102.L10_Text <> vbNullString Then 'Berat belian keseluruhan (Barang trade in)
                rs!berat_belian = Format(Frm102.L10_Text, "0.00")
            Else
                rs!berat_belian = "0.00"
            End If
            If Frm102.L11_Text <> vbNullString Then 'Beza antara berat jualan dan belian
                rs!beza_berat = Format(Frm102.L11_Text, "0.00")
            Else
                rs!beza_berat = "0.00"
            End If
            If Frm102.TB11 <> vbNullString Then 'Harga semasa (penilaian harga emas oleh pihak kedai)
                rs!harga_Semasa = Format(Frm102.TB11, "0.00")
            Else
                rs!harga_Semasa = "0.00"
            End If
            If Frm102.L12_Text <> vbNullString Then 'Nilaian harga emas oleh pihak kedai terhadap beza berat tersebut (jika bayaran perlu dibuat oleh pihak kedai sahaja)
                rs!harga_emas = Format(Frm102.L12_Text, "0.00")
            Else
                rs!harga_emas = "0.00"
            End If
            If Frm102.L31_Text <> vbNullString Then 'Harga emas tanpa GST
                rs!harga_tanpa_gst = Format(Frm102.L31_Text, "0.00")
            Else
                rs!harga_tanpa_gst = "0.00"
            End If
            
            If Frm102.CB5 = 1 Then
                rs!gst_ari_nashi = "ZR (L)" '0 : Tiada GST : ZR(L) , 1 : Ada GST : SR
                rs!gst_include = 0 '0 : GST Dibayar Oleh Pelanggan , 1 : GST Termasuk Dalam Harga Barang
            ElseIf Frm102.CB6 = 1 Then
                rs!gst_ari_nashi = "SR" '0 : Tiada GST : ZR(L) , 1 : Ada GST : SR
                
                If Frm102.CB7 = 1 Then 'Jenis Cukai GST SR
                    rs!gst_include = 1 '0 : GST Dibayar Oleh Pelanggan , 1 : GST Termasuk Dalam Harga Barang
                Else
                    rs!gst_include = 0 '0 : GST Dibayar Oleh Pelanggan , 1 : GST Termasuk Dalam Harga Barang
                End If
            End If
            If Frm102.L21_Text <> vbNullString Then
                rs!kadar_gst = Frm102.L21_Text 'Kadar Cukai GST (%)
            Else
                rs!kadar_gst = "0" 'Jumlah Cukai GST (RM)
            End If
            If Frm102.L31_Text <> vbNullString Then 'Harga emas tanpa GST
                rs!harga_tanpa_gst = Format(Frm102.L31_Text, "0.00")
            Else
                rs!harga_tanpa_gst = "0.00"
            End If
            If Frm102.TB12 <> vbNullString Then 'Jumlah Cukai GST (RM)
                rs!jumlah_gst = Format(Frm102.TB12, "0.00")
            Else
                rs!jumlah_gst = "0.00"
            End If
            If Frm102.TB13 <> vbNullString Then 'Jumlah emas + GST (RM)
                rs!harga_dengan_gst = Format(Frm102.TB13, "0.00")
            Else
                rs!harga_dengan_gst = "0.00"
            End If
            If Frm102.CBB4 <> vbNullString Then
                Frm102_LM_EMP_NO = Split(Frm102.CBB4, "  |  ")(1)
                rs!no_pekerja = Frm102_LM_EMP_NO 'No. Pekerja
            End If
            rs!write_timestamp2 = Now

            rs.Update
        End If
        
        rs.Close
        Set rs = Nothing
'### Masukkan data voucher / invoice bagi belian agen ini ### - End
    
        If DATA_SAVE = 1 Then
    '###Update No. Resit### - Start
            G_No_RESIT_JUALAN = vbNullString
            G_No_RESIT_JUALAN = Frm102.L23_Text
            
    '#### Update Log Aktiviti Sistem #### - Start
            If Frm102.CBB4 <> vbNullString Then
                Frm102_LM_EMP_NAME = Split(Frm102.CBB4, "  |  ")(0)
            End If
        
            'User = MDI_frm1.L3_Text
            LogAct_Memory = "[" & Frm102_LM_EMP_NAME & "] Edit jualan barang kemas kepada agen. No. Invoice [" & Frm102.L23_Text & "]."
            LogDate_Memory = DateTime.Date & " " & DateTime.Time$
            Call UpdateLog_Database
    '#### Update Log Aktiviti Sistem #### - End
            
            If Frm102_LM_GENERATED = 1 Then '0 : Tiada No Voucher yang dihasilkan , 1 : Ada No. Voucher yang dihasilkan
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from default_setting", cn, adOpenKeyset, adLockOptimistic
                
                If Not rs.EOF Then
                    If rs!Default1 = "Default" Then
                        If Frm102.L34_Text.Visible = True Then
                            rs!no_trade_in_agen = Frm102.L22_Text + 1 'No. Voucher Trade In
                        End If
                        rs.Update
                    End If
                End If
                
                rs.Close
                Set rs = Nothing
            End If
            
            Note = "Data Telah Berjaya Disimpan." & vbCrLf & _
                    "Refresh Data Anda ?"

            Answer = MsgBox(Note, vbQuestion + vbOK, "Confirmation")
            
            If Answer = vbOK Then
                GM_NEXT_PREV = 2
                
                If Frm101.L33_Text = 0 Then '0 : Carian Report Biasa ,  1 : Carian Ikut Berat , 2 : Carian Ikut No. Resit Jualan , 3 : Carian Ikut No. Resit Buyback / Trade In , 4 : Carian mengikut No. Invoice Supplier , 5 : Carian mengikut No. Siri Produk (Belian BK) , 6 : Carian mengikut No. Siri Produk (Buyback BK) , 7 : Carian mengikut No. Siri Produk (Belian GB) , 8 : Carian mengikut No. Siri Produk (Buyback GB)
                    If Frm101.CB3 = 1 Then 'Report Jualan
                        Call Frm85_Header_Report_Jualan
                        Call Frm85_Report_Jualan_page
                    End If
                ElseIf Frm101.L33_Text = 2 Then '0 : Carian Report Biasa ,  1 : Carian Ikut Berat , 2 : Carian Ikut No. Resit Jualan , 3 : Carian Ikut No. Resit Buyback / Trade In , 4 : Carian mengikut No. Invoice Supplier , 5 : Carian mengikut No. Siri Produk (Belian BK) , 6 : Carian mengikut No. Siri Produk (Buyback BK) , 7 : Carian mengikut No. Siri Produk (Belian GB) , 8 : Carian mengikut No. Siri Produk (Buyback GB)
                    Call Frm85_Header_Report_Jualan
                    Call Frm85_carian_jualan_page
                ElseIf Frm101.L33_Text = 5 Then '0 : Carian Report Biasa ,  1 : Carian Ikut Berat , 2 : Carian Ikut No. Resit Jualan , 3 : Carian Ikut No. Resit Buyback / Trade In , 4 : Carian mengikut No. Invoice Supplier , 5 : Carian mengikut No. Siri Produk (Belian BK) , 6 : Carian mengikut No. Siri Produk (Buyback BK) , 7 : Carian mengikut No. Siri Produk (Belian GB) , 8 : Carian mengikut No. Siri Produk (Buyback GB)
                    Call Frm85_Header_Report_Jualan
                    Call Frm85_Report_Jualan_barcode
                End If
                
                Frm85.Show
                Unload Frm102
                MDI_frm1.L5_Text = 12
            Else
                Frm85.Show
                Unload Frm102
                MDI_frm1.L5_Text = 12
            End If
            
            MsgBox "Data Telah Berjaya Disimpan.", vbInformation, "Info"
    '###Update No. Resit### - End
        End If
        
    End If
End If
End Sub
Private Sub CMD11_Click()
'On Error Resume Next
Frm85.Show
Unload Frm102
MDI_frm1.L5_Text = 12
End Sub
Private Sub CMD12_Click()
'On Error Resume Next
If Frm102.L46_Text = vbNullString Then

    Call Frm28_initial
    
    Frm28.Show 1
    
Else
    Frm28.Show 1
End If
End Sub
Private Sub CMD13_Click()
'On Error Resume Next
Frm102.Pic1.Visible = False
End Sub
Private Sub CMD2_Click()
'On Error Resume Next
Dim Err(30)
Dim Frm102_LM_BERAT_ASAL As Double
Dim Frm102_LM_BERAT_JUAL As Double
Dim Frm102_LM_HARGA_MODAL As Double
Dim Frm102_LM_HARGA_JUAL As Double
Dim Frm102_LM_HARGA_SEMASA_MODAL As Double
Dim Frm102_LM_TETAPANHARGA As Double
Dim Frm102_LM_LIMIT As Double
Dim Frm102_LM_HARGA_STAFF As Double 'Tetapan harga jualan kepada staff
Dim Frm102_LM_HARGA_PELANGGAN As Double 'Tetapan harga jualan kepada pelanggan
Dim Frm102_LM_BERAT_JUAL_ASAL As Double 'Berat Jualan (Purity Asal)
Dim Frm102_LM_HARGA_SEMASA_999 As Double 'Harga semasa (jualan) (Purity 999.9)
Dim Frm102_LM_HARGA_SUPPLIER As Double 'Harga per gram (harga semasa) dari supplier (modal)
Dim Frm102_LM_BERAT_999 As Double 'Berat Jualan (Purity Asal)
Dim Frm102_UPAH_MODAL As Double 'Upah modal
Dim Frm102_UPAH_JUAL As Double 'Upah jualan

Frm102_UPAH_MODAL = 0 'Upah modal
Frm102_UPAH_JUAL = 0 'Upah jualan
Frm102_LM_BERAT_JUAL_ASAL = 0 'Berat Jualan (Purity Asal)
Frm102_LM_HARGA_SEMASA_999 = 0 'Harga semasa (jualan) (Purity 999.9)
Frm102_LM_HARGA_SUPPLIER = 0 'Harga per gram (harga semasa) dari supplier (modal)
Frm102_LM_BERAT_999 = 0 'Berat Jualan (Purity Asal)
x = 0
Frm102_LM_BERAT_ASAL = 0
Frm102_LM_BERAT_JUAL = 0
Frm102_LM_DATA_SAVE = 0
Frm102_LM_HARGA_MODAL = 0
Frm102_LM_HARGA_JUAL = 0
Frm102_LM_HARGA_SEMASA_MODAL = 0
Frm102_LM_PRICE_CHECK = 0 '0 : Tidak Perlu Periksa Harga Semasa Jualan , 1 : Perlu Periksa Harga Semasa Jualan
Frm102_LM_TETAPANHARGA = 0
Frm102_LM_LIMIT = 0
Frm102_LM_HARGA_STAFF = 0
Frm102_LM_HARGA_PELANGGAN = 0

If Frm102.L3_Text = vbNullString Then
    x = x + 1
    Err(x) = "Sila masukkan [No. Siri Produk]."
End If
If Frm102.L33_Text = vbNullString Then
    x = x + 1
    Err(x) = "Maklumat harga semasa modal belian item ini yang tidak sah. Sila keluar dari menu ini dan cuba sekali lagi."
End If
If Frm102.L50_Text = vbNullString Or (Frm102.L50_Text <> vbNullString And Not IsNumeric(Frm102.L50_Text)) Then
    x = x + 1
    Err(x) = "Maklumat upah modal yang tidak sah. Sila keluar dari menu ini dan cuba sekali lagi."
End If
If Frm102.L6_Text = vbNullString Or (Frm102.L6_Text <> vbNullString And Not IsNumeric(Frm102.L6_Text)) Then
    x = x + 1
    Err(x) = "Sila maklumat [Berat Asal]. Sila scan item sekali lagi."
End If
If Frm102.TB3 = vbNullString Or (Frm102.TB3 <> vbNullString And Not IsNumeric(Frm102.TB3)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Berat Jualan]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm102.TB2 = vbNullString Or (Frm102.TB2 <> vbNullString And Not IsNumeric(Frm102.TB2)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Harga Semasa Emas 999.9]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm102.TB2 <> vbNullString And IsNumeric(Frm102.TB2) Then

    If Format(Frm102.TB2, "0.00") = "0.00" Then
        x = x + 1
        Err(x) = "Harga emas semasa 999.9 yang tidak sah. Nilai 0.00 tidak dibenarkan."
    End If
    
End If
If Frm102.CBB1 = vbNullString Then
    x = x + 1
    Err(x) = "Sila pilih [Kadar Tukaran Purity 999.9]."
End If
If Frm102.L7_Text = vbNullString Or (Frm102.L7_Text <> vbNullString And Not IsNumeric(Frm102.L7_Text)) Then
    x = x + 1
    Err(x) = "[Berat 999.9] yang tidak sah. Sila scan item sekali lagi."
End If
If Frm102.TB4 = vbNullString Or (Frm102.TB4 <> vbNullString And Not IsNumeric(Frm102.TB4)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Upah]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm102.CB2 = 0 And Frm102.CB3 = 0 Then
    x = x + 1
    Err(x) = "Sila buat pilihan jenis GST bagi upah."
End If
If Frm102.TB5 = vbNullString Or Frm102.TB6 = vbNullString Then
    x = x + 1
    Err(x) = "Maklumat berkenaan GST yang tidak sah. Sila keluar dari menu ini dan cuba sekali lagi."
End If

If (Frm102.L6_Text <> vbNullString And IsNumeric(Frm102.L6_Text)) And (Frm102.TB3 <> vbNullString And IsNumeric(Frm102.TB3)) Then
    Frm102_LM_BERAT_ASAL = Frm102.L6_Text 'Berat Asal
    Frm102_LM_BERAT_JUAL = Frm102.TB3 'Berat Jualan
    
    If Frm102_LM_BERAT_JUAL > Frm102_LM_BERAT_ASAL Then
        x = x + 1
        Err(x) = "Berat jualan melebihi berat asal."
    End If
End If
If Frm102.L49_Text = vbNullString Or (Frm102.L49_Text <> vbNullString And Not IsNumeric(Frm102.L49_Text)) Then
    x = x + 1
    Err(x) = "Tiada maklumat harga dari supplier bagi purity ini." & vbCrLf & _
                "Sila pastikan harga dari supplier bagi purity ini telah ditetapkan dalam TETAPAN HARIAN SISTEM."
End If

If x <> 0 Then
    Frm6.Show
    Frm6.Pic1.Cls
    For Y = 1 To x
        Frm6.Pic1.Print Y & " - " & Err(Y)
    Next Y
    Exit Sub
Else
    Note = "Adakah anda ingin masukkan item ini ke dalam senarai jualan ?"
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbNo Then
        Exit Sub
    End If
    If Answer = vbYes Then
    
'### Periksa Data Dulang ### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from Data_Database where no_siri_Produk='" & Frm102.L3_Text & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Not IsNull(rs!dulang) Then Frm102_LM_DULANG = rs!dulang 'Dulang
        End If
        
        rs.Close
        Set rs = Nothing
'### Periksa Data Dulang ### - End
        
'### Masukkan Data Ke Dalam Temp Table ### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from " & G_JUALAN_TEMP & " where ID='" & Frm102.L24_Text & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Frm102.L3_Text <> vbNullString Then
                rs!no_siri_Produk = Frm102.L3_Text 'No. Siri Produk
            Else
                rs!no_siri_Produk = Null 'No. Siri Produk
            End If
            If Frm102.L5_Text <> vbNullString Then
                rs!kategori_Produk = Frm102.L5_Text 'Kategori Produk
            Else
                rs!kategori_Produk = Null 'Kategori Produk
            End If
            If Frm102.L4_Text <> vbNullString Then
                rs!purity = Frm102.L4_Text 'Purity
            Else
                rs!purity = Null 'Purity
            End If
            If Frm102.L6_Text <> vbNullString Then
                rs!Berat_Asal = Format(Frm102.L6_Text, "0.00") 'Berat Asal (g)
            Else
                rs!Berat_Asal = Null 'Berat Asal (g)
            End If
            If Frm102.TB3 <> vbNullString Then
                rs!berat_jualan = Format(Frm102.TB3, "0.00") 'Berat Jualan (g)
            Else
                rs!berat_jualan = Null 'Berat Jualan (g)
            End If
            If Frm102.TB2 <> vbNullString Then
                rs!harga_Semasa = Format(Frm102.TB2, "0.00") 'Harga Semasa (RM/g)
            Else
                rs!harga_Semasa = Null 'Harga Semasa (RM/g)
            End If
            If Frm102.TB4 <> vbNullString Then
                rs!UPAH = Format(Frm102.TB4, "0.00") 'Upah (RM)
            Else
                rs!UPAH = Null 'Upah (RM)
            End If
            
            Frm102_LM_HARGA_SEMASA = Frm102.TB2 'Harga emas semasa 999.9 (Untuk tujuan jualan kepada pelanggan)
            Frm102_LM_BERAT_JUALAN_9999 = Frm102.L7_Text 'Berat jualan dalam purity 999.9
            Frm102_LM_UPAH_DAN_GST = Frm102.TB6 'Jumlah Upah + GST (Bagi jualan setiap item)

            If Frm102.TB6 <> vbNullString Then
                rs!harga_asal = Format((Frm102_LM_HARGA_SEMASA * Frm102_LM_BERAT_JUALAN_9999) + Frm102_LM_UPAH_DAN_GST, "0.00") 'Harga Asal Item (RM)
            Else
                rs!harga_asal = Null 'Harga Asal Item (RM)
            End If
            
            rs!diskaun = "0.00" 'Diskaun (%)
            rs!harga_lepas_diskaun = Format((Frm102_LM_HARGA_SEMASA * Frm102_LM_BERAT_JUALAN_9999) + Frm102_LM_UPAH_DAN_GST, "0.00") 'Harga Asal Item (RM)
            rs!adjustment = Format(0, "0.00") 'Adjustment (RM)
            rs!harga_jualan = Format((Frm102_LM_HARGA_SEMASA * Frm102_LM_BERAT_JUALAN_9999) + Frm102_LM_UPAH_DAN_GST, "0.00") 'Harga Asal Item (RM)
            rs!harga_jualan_dengan_gst = Format((Frm102_LM_HARGA_SEMASA * Frm102_LM_BERAT_JUALAN_9999) + Frm102_LM_UPAH_DAN_GST, "0.00") 'Harga Asal Item (RM)
            
            If Frm102.CB2 = 1 Then
                rs!gst_ari_nashi = "ZR (L)" '0 : Tiada GST : ZR(L) , 1 : Ada GST : SR
                rs!kadar_gst = Null 'Kadar Cukai GST (%)
                If Frm102.TB5 <> vbNullString Then
                    rs!jumlah_gst = Format(Frm102.TB5, "0.00") 'Jumlah GST (Bagi jualan setiap item)
                Else
                    rs!jumlah_gst = "0.00" 'Jumlah GST (Bagi jualan setiap item)
                End If
            ElseIf Frm102.CB3 = 1 Then
                rs!gst_ari_nashi = "SR" '0 : Tiada GST : ZR(L) , 1 : Ada GST : SR
                If Frm102.L21_Text <> vbNullString Then
                    rs!kadar_gst = Frm102.L21_Text 'Kadar Cukai GST (%)
                Else
                    rs!kadar_gst = "0" 'Jumlah Cukai GST (RM)
                End If
                If Frm102.TB5 <> vbNullString Then
                    rs!jumlah_gst = Format(Frm102.TB5, "0.00") 'Jumlah GST (Bagi jualan setiap item)
                Else
                    rs!jumlah_gst = "0.00" 'Jumlah GST (Bagi jualan setiap item)
                End If
                If Frm102.CB4 = 1 Then 'Jenis Cukai GST SR
                    rs!gst_include = 1 '0 : GST Dibayar Oleh Pelanggan , 1 : GST Termasuk Dalam Harga Barang
                Else
                    rs!gst_include = 0 '0 : GST Dibayar Oleh Pelanggan , 1 : GST Termasuk Dalam Harga Barang
                End If
            End If
            If Frm102.L30_Text <> vbNullString Then
                rs!harga_tanpa_gst = Format(Frm102.L30_Text, "0.00") 'Harga Jualan Tanpa GST (RM)
            Else
                rs!harga_tanpa_gst = Null 'Harga Jualan Tanpa GST (RM)
            End If
            If Frm102.TB6 <> vbNullString Then
                rs!harga_dengan_gst = Format(Frm102.TB6, "0.00") 'Harga Jualan Termasuk GST (RM)
            Else
                rs!harga_dengan_gst = Null 'Harga Jualan Termasuk GST (RM)
            End If
            rs!dropship = 0 '0 : Jualan Bukan Oleh Agen Dropship , 1 : Jualan Oleh Agen Dropship
            rs!komisyen_per_gram = Null 'Komisyen Per Gram Dropship (RM/g)
            rs!jumlah_komisyen = Null 'Jumlah Komisyen Kepada Agen Dropship Bagi Item Ini
            
'Status
'0 : Keluarkan Dari Senarai
'1 : Data Baru (Fresh)
'2 : Data Baru Diedit (Fresh)
'3 : Data Baru Dari Menu Edit
'4 : Data Baru Dari Menu Edit Yang Telah Diedit

            If Frm102.L32_Text = "0" Then '0 : Menu Data Baru , 1 : Menu Edit Data
                rs!Status = 1
            ElseIf Frm102.L32_Text = "1" Then
                If rs!Status = "2" Then
                    rs!Status = 3
                End If
                If rs!Status = "4" Then
                    rs!Status = 4
                End If
            End If
            
            rs!Type = 0 '0 : BK , 1 : Barang Permata
            If Frm102.L33_Text <> vbNullString Then
                rs!harga_per_gram_modal = Format(Frm102.L33_Text, "0.00") 'Harga Per Gram Bagi Modal (RM/g)
                Frm102_LM_HARGA_SEMASA_MODAL = Frm102.L33_Text
            Else
                rs!harga_per_gram_modal = Format(0, "0.00") 'Harga Per Gram Bagi Modal (RM/g)
            End If
            rs!modal = Format(Frm102_LM_HARGA_SEMASA_MODAL * Frm102_LM_BERAT_JUALAN_9999, "0.00") 'Harga Modal (RM)
            If IsNumeric(Frm102.TB6) And IsNumeric(Frm102.L33_Text) And IsNumeric(Frm102.TB3) Then
                Frm102_LM_HARGA_MODAL = Frm102.L33_Text * Frm102.TB3 'Harga modal
                Frm102_LM_HARGA_JUAL = (Frm102_LM_HARGA_SEMASA * Frm102_LM_BERAT_JUALAN_9999) + Frm102_LM_UPAH_DAN_GST 'Harga jualan
                
                rs!untung = Format(Frm102_LM_HARGA_JUAL - Frm102_LM_HARGA_MODAL, "0.00") 'Jumlah Keuntungan
            Else
                rs!untung = Format(0, "0.00") 'Jumlah Keuntungan
            End If

            If Frm102.L49_Text <> vbNullString Then 'Harga per gram (harga semasa) dari supplier (modal)
                rs!harga_per_gram_supplier = Frm102.L49_Text
            Else
                rs!harga_per_gram_supplier = 0
            End If
            
            If Frm102.L49_Text <> vbNullString Then 'Harga per gram (harga semasa) dari supplier (modal)
                rs!harga_per_gram_supplier = Frm102.L49_Text
            Else
                rs!harga_per_gram_supplier = 0
            End If
            
            If IsNumeric(Frm102.TB3) And IsNumeric(Frm102.TB2) And IsNumeric(Frm102.L49_Text) And IsNumeric(Frm102.L7_Text) And IsNumeric(Frm102.L6_Text) And IsNumeric(Frm102.L50_Text) And IsNumeric(Frm102.TB4) Then
                Frm102_LM_BERAT_JUAL_ASAL = Frm102.TB3 'Berat Jualan (Purity Asal)
                Frm102_LM_BERAT_ASAL = Frm102.L6_Text 'Berat Asal (Purity Asal)
                Frm102_UPAH_JUAL = Frm102.TB4 'Upah jualan
                Frm102_UPAH_MODAL = Frm102.L50_Text 'Upah modal
                Frm102_LM_HARGA_SEMASA_999 = Frm102.TB2 'Harga semasa (jualan) (Purity 999.9)
                Frm102_LM_HARGA_SUPPLIER = Frm102.L49_Text 'Harga per gram (harga semasa) dari supplier (modal)
                Frm102_LM_BERAT_999 = Frm102.L7_Text 'Berat emas dalam purity 999.9
                
                rs!upah_modal = Frm102.L50_Text 'Upah modal
                rs!harga_per_gram_supplier = Frm102.L49_Text 'Harga per gram (harga semasa) dari supplier (modal)
                rs!untung2 = Format(((Frm102_LM_BERAT_999 * Frm102_LM_HARGA_SEMASA_999) + Frm102_UPAH_JUAL) - ((Frm102_LM_BERAT_JUAL_ASAL * Frm102_LM_HARGA_SUPPLIER) + (Frm102_LM_BERAT_JUAL_ASAL / Frm102_LM_BERAT_ASAL) * Frm102_UPAH_MODAL), "0.00") 'Untung jika restok pada harga supplier ini
                
            Else
            
                rs!harga_per_gram_supplier = "0.00" 'Harga per gram (harga semasa) dari supplier (modal)
                rs!untung2 = "0.00" 'Untung jika restok pada harga supplier ini
                rs!upah_modal = "0.00" 'Upah modal
                
            End If
            
            If Format(Frm102.L6_Text, "0.00") = Format(Frm102.TB3, "0.00") Then
                rs!potong_flag = 0 '0 : Tiada Potong , 1 : Ada Potong
            Else
                rs!potong_flag = 1 '0 : Tiada Potong , 1 : Ada Potong
            End If
            rs!dulang = Frm102_LM_DULANG 'Dulang
            If Frm102.CBB1 <> vbNullString Then
                rs!pemalar_tukaran_999 = Frm102.CBB1 'Pemalar tukaran emas kepada 999.9 bagi urusan jualan
            Else
                rs!pemalar_tukaran_999 = Null 'Pemalar tukaran emas kepada 999.9 bagi urusan jualan
            End If
            If Frm102.L7_Text <> vbNullString Then
                rs!berat_999 = Format(Frm102.L7_Text, "0.00") 'Berat jualan dalam purity 999.9
            Else
                rs!berat_999 = Null 'Berat jualan dalam purity 999.9
            End If
            
            rs.Update

            Frm102_LM_DATA_SAVE = 1
        End If
        
        rs.Close
        Set rs = Nothing
'### Masukkan Data Ke Dalam Temp Table ### - End
        
        If Frm102_LM_DATA_SAVE = 1 Then
            Call frm102_reset_1
            Call Frm102_Senarai_Jualan_Header
            Call Frm102_Senarai_Jualan
            
            MsgBox "Data yang telah diedit telah berjaya dimasukkan ke dalam senarai jualan.", vbInformation, "Info"
            
            Frm102.TB1.SetFocus
        End If
    End If
End If
End Sub
Private Sub CMD3_Click()
'On Error Resume Next
Note = "Adakah anda ingin batalkan urusan edit data ini ?"

Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")

If Answer = vbNo Then
    Exit Sub
End If
If Answer = vbYes Then

    Call frm102_reset_1
    
    Frm102.CMD1.Visible = True
    Frm102.CMD2.Visible = False
    Frm102.CMD3.Visible = False
    
    Frm102.TB1.SetFocus
    
End If
End Sub
Private Sub CMD4_Click()
'On Error Resume Next
Dim Err(10)

x = 0

If Frm102.TB10 = vbNullString Or (Frm102.TB10 <> vbNullString And Not IsNumeric(Frm102.TB10)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Berat]. Hanya NOMBOR dibenarkan di dalam ruangan ini."
End If
If Frm102.TB10 <> vbNullString And IsNumeric(Frm102.TB10) Then
    If Format(Frm102.TB10, "0.00") = "0.00" Then
        x = x + 1
        Err(x) = "Berat yang tidak sah. Hanya NOMBOR dibenarkan dalam ruangan ini dan nilai adalah 0.00 tidak dibenarkan."
    End If
End If
If Frm102.TB24 <> vbNullString And IsNumeric(Frm102.TB24) Then
    If Format(Frm102.TB24, "0.00") = "0.00" Then
        x = x + 1
        Err(x) = "Kadar tukaran purity 999.9 yang tidak sah. Hanya NOMBOR dibenarkan dalam ruangan ini dan nilai adalah 0.00 tidak dibenarkan."
    End If
End If
If Frm102.CBB2 = vbNullString Then
    x = x + 1
    Err(x) = "Sila pilih [Purity Barang]."
End If
If Frm102.TB24 = vbNullString Or (Frm102.TB24 <> vbNullString And Not IsNumeric(Frm102.TB24)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Kadar Tukaran Purity 999.9]. Hanya NOMBOR dibenarkan di dalam ruangan ini."
End If
If Frm102.L47_Text = vbNullString Then
    x = x + 1
    Err(x) = "Tiada maklumat kod purity. Sila cuba buat pilihan purity sekali lagi dan simpan data sekali lagi."
End If

If x <> 0 Then
    Frm6.Show
    Frm6.Pic1.Cls
    For Y = 1 To x
        Frm6.Pic1.Print Y & " - " & Err(Y)
    Next Y
    Exit Sub
Else

    Note = "Adakah anda ingin masukkan item ini ke dalam senarai belian ?"
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbNo Then
        Exit Sub
    End If
    If Answer = vbYes Then
    
'### Masukkan Data Ke Dalam Temp Table ### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 49_belian_temp", cn, adOpenKeyset, adLockOptimistic
        
        rs.AddNew
        If Frm102.TB10 <> vbNullString Then 'Berat (barang trade in)
            rs!Berat_Asal = Format(Frm102.TB10, "0.00")
        Else
            rs!Berat_Asal = Null
        End If
        If Frm102.CBB2 <> vbNullString Then 'Trade In : Purity barang
            rs!purity = Frm102.CBB2
        Else
            rs!purity = Null
        End If
        If Frm102.L47_Text <> vbNullString Then 'Kod purity
            rs!kod_Purity = Frm102.L47_Text
        Else
            rs!kod_Purity = Null
        End If
        If Frm102.TB24 <> vbNullString Then 'Trade In : Kadar tukaran purity 999.9
            rs!kadar_tukaran = Frm102.TB24
        Else
            rs!kadar_tukaran = Null
        End If
        If Frm102.L8_Text <> vbNullString Then 'Trade In : Berat dalam 999.9
            rs!berat_tukaran = Format(Frm102.L8_Text, "0.00")
        Else
            rs!berat_tukaran = Null
        End If
        
'Status
'1 : Kemasukkan data baru (Data baru)
'2 : Tiada perubahan (Menu edit)
'3 : Data telah diedit (Menu edit)
'4 : Kemasukkan data baru (Menu edit)
'5 : Data dipadamkan (Data yang diterima dari menu baru)

        If Frm102.L32_Text = 0 Then '0 : Data Baru , 1 : Edit Data
            rs!Status = 1
        ElseIf Frm102.L32_Text = 1 Then
            rs!Status = 4
        End If
        
        rs.Update
        
        rs.Close
        Set rs = Nothing
'### Masukkan Data Ke Dalam Temp Table ### - End

        Call frm102_reset_2
        Call Frm102_senarai_belian_header
        Call Frm102_senarai_belian
        
        Frm102.TB10.SetFocus

    End If
    
End If
End Sub
Private Sub CMD5_Click()
'On Error Resume Next
Dim Err(10)

x = 0

If Frm102.TB10 = vbNullString Or (Frm102.TB10 <> vbNullString And Not IsNumeric(Frm102.TB10)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Berat]. Hanya NOMBOR dibenarkan di dalam ruangan ini."
End If
If Frm102.TB10 <> vbNullString And IsNumeric(Frm102.TB10) Then
    If Format(Frm102.TB10, "0.00") = "0.00" Then
        x = x + 1
        Err(x) = "Berat yang tidak sah. Hanya NOMBOR dibenarkan dalam ruangan ini dan nilai adalah 0.00 tidak dibenarkan."
    End If
End If
If Frm102.TB24 <> vbNullString And IsNumeric(Frm102.TB24) Then
    If Format(Frm102.TB24, "0.00") = "0.00" Then
        x = x + 1
        Err(x) = "Kadar tukaran purity 999.9 yang tidak sah. Hanya NOMBOR dibenarkan dalam ruangan ini dan nilai adalah 0.00 tidak dibenarkan."
    End If
End If
If Frm102.CBB2 = vbNullString Then
    x = x + 1
    Err(x) = "Sila pilih [Purity Barang]."
End If
If Frm102.TB24 = vbNullString Or (Frm102.TB24 <> vbNullString And Not IsNumeric(Frm102.TB24)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Kadar Tukaran Purity 999.9]. Hanya NOMBOR dibenarkan di dalam ruangan ini."
End If
If Frm102.L47_Text = vbNullString Then
    x = x + 1
    Err(x) = "Tiada maklumat kod purity. Sila cuba buat pilihan purity sekali lagi dan simpan data sekali lagi."
End If

If x <> 0 Then
    Frm6.Show
    Frm6.Pic1.Cls
    For Y = 1 To x
        Frm6.Pic1.Print Y & " - " & Err(Y)
    Next Y
    Exit Sub
Else

    Note = "Adakah anda ingin masukkan data yang telah diedit ke dalam senarai belian ?"
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbNo Then
        Exit Sub
    End If
    If Answer = vbYes Then
    
'### Masukkan Data Ke Dalam Temp Table ### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 49_belian_temp where ID='" & Frm102.L25_Text & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Frm102.TB10 <> vbNullString Then 'Berat (barang trade in)
                rs!Berat_Asal = Format(Frm102.TB10, "0.00")
            Else
                rs!Berat_Asal = Null
            End If
            If Frm102.CBB2 <> vbNullString Then 'Trade In : Purity barang
                rs!purity = Frm102.CBB2
            Else
                rs!purity = Null
            End If
            If Frm102.L47_Text <> vbNullString Then 'Kod purity
                rs!kod_Purity = Frm102.L47_Text
            Else
                rs!kod_Purity = Null
            End If
            If Frm102.TB24 <> vbNullString Then 'Trade In : Kadar tukaran purity 999.9
                rs!kadar_tukaran = Frm102.TB24
            Else
                rs!kadar_tukaran = Null
            End If
            If Frm102.L8_Text <> vbNullString Then 'Trade In : Berat dalam 999.9
                rs!berat_tukaran = Format(Frm102.L8_Text, "0.00")
            Else
                rs!berat_tukaran = Null
            End If
            
'Status
'1 : Kemasukkan data baru (Data baru)
'2 : Tiada perubahan (Menu edit)
'3 : Data telah diedit (Menu edit)
'4 : Kemasukkan data baru (Menu edit)
'5 : Data dipadamkan (Data yang diterima dari menu baru)

            If Frm102.L32_Text = 0 Then '0 : Data Baru , 1 : Edit Data
                rs!Status = 1
            ElseIf Frm102.L32_Text = 1 Then
                If rs!Status = 2 Then
                    rs!Status = 3
                ElseIf rs!Status = 4 Then
                    rs!Status = 4
                End If
            End If
            
            rs.Update
        End If
        
        rs.Close
        Set rs = Nothing
'### Masukkan Data Ke Dalam Temp Table ### - End

        Call frm102_reset_2
        Call Frm102_senarai_belian_header
        Call Frm102_senarai_belian

    End If
    
End If
End Sub
Private Sub CMD6_Click()
'On Error Resume Next

Note = "Adakah anda ingin batalkan urusan edit data ini ?"

Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")

If Answer = vbNo Then
    Exit Sub
End If
If Answer = vbYes Then

    Call frm102_reset_2
    
    Frm102.CMD4.Visible = True 'Masukkan dalam senarai trade in
    Frm102.CMD5.Visible = False 'Masukkan dalam senarai trade in (Edit)
    Frm102.CMD6.Visible = False 'Batal edit data
    
    Frm102.TB10.SetFocus
    
End If
End Sub
Private Sub CMD7_Click()
'On Error Resume Next
If Frm102.TB1 = vbNullString Then

    MsgBox "Sila masukkan No. Siri Produk.", vbExclamation, "Error"
    
    Frm102.TB1.SetFocus
    Exit Sub

End If

If InStr(1, Frm102.TB1, "'") <> 0 Then

    MsgBox "No. Siri Produk Mengandungi Simbol Yang Tidak Sah , ['].", vbInformation, "Info"
    
    Frm102.TB1 = vbNullString
    
    Frm102.TB1.SetFocus
    Exit Sub
    
End If

Call frm102_reset_1
Call Frm102_Call_Product_Detail
End Sub
Private Sub CMD8_Click()
'On Error Resume Next
Dim Err(30)
Dim rs1 As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim Frm102_LM_CUKAI_ZR As Double
Dim Frm102_LM_CUKAI_SR As Double
Dim Frm102_LM_BERAT_ASAL As Double
Dim Frm102_LM_BERAT_JUALAN As Double

Frm102_LM_KATEGORI = 0
Frm102_LM_BERAT_ASAL = 0 'Berat Asal (g)
Frm102_LM_BERAT_JUALAN = 0 'Berat Jualan (g)
Frm102_LM_CUKAI_ZR = 0 'Jumlah cukai GST ZR
Frm102_LM_CUKAI_SR = 0 'Jumlah cukai GST SR

Frm102_LM_Flag_SIMPANAN = 0 '0 : Tiada Penggunakan Duit Simpanan Kedai , 1 : Ada Penggunakan Duit Simpanan Kedai

If Frm102.TB14 = vbNullString Or (Frm102.TB14 <> vbNullString And Not IsNumeric(Frm102.TB14)) Then
    x = x + 1
    Err(x) = "Hanya NOMBOR dibenarkan dalam ruangan bayaran secara tunai. Sila masukkan 0 jika tiada bayaran tunai."
End If
If Frm102.TB15 = vbNullString Or (Frm102.TB15 <> vbNullString And Not IsNumeric(Frm102.TB15)) Then
    x = x + 1
    Err(x) = "Hanya NOMBOR dibenarkan dalam ruangan bayaran secara bank in. Sila masukkan 0 jika tiada bayaran bank in."
End If
If Frm102.TB16 = vbNullString Or (Frm102.TB16 <> vbNullString And Not IsNumeric(Frm102.TB16)) Then
    x = x + 1
    Err(x) = "Hanya NOMBOR dibenarkan dalam ruangan bayaran secara kad kredit. Sila masukkan 0 jika tiada bayaran kad kredit."
End If
If Frm102.TB22 = vbNullString Or (Frm102.TB22 <> vbNullString And Not IsNumeric(Frm102.TB22)) Then
    x = x + 1
    Err(x) = "Hanya NOMBOR dibenarkan dalam ruangan bayaran secara duit simpanan di kedai. Sila masukkan 0 jika tiada bayaran simpanan di kedai."
End If
If Frm102.TB19 = vbNullString Or (Frm102.TB19 <> vbNullString And Not IsNumeric(Frm102.TB19)) Then
    x = x + 1
    Err(x) = "Hanya NOMBOR dibenarkan dalam ruangan bayaran secara debit kad. Sila masukkan 0 jika tiada bayaran debit kad."
End If
If Format(Frm102.L14_Text, "0.00") <> Format(Frm102.TB23, "0.00") Then
    x = x + 1
    Err(x) = "Jumlah bayaran tidak sama dengan jumlah yang perlu dibayar."
End If
If Frm102.L43_Text = 0 Then
    x = x + 1
    Err(x) = "Tiada senarai jualan."
End If
If Frm102.L44_Text = 0 Then
    x = x + 1
    Err(x) = "Tiada senarai barang trade in."
End If
If Frm102.CBB4 = vbNullString Then
    x = x + 1
    Err(x) = "Sila pilih nama pekerja."
End If
'If Frm102.L46_Text = vbNullString Then
'    X = X + 1
'    Err(X) = "Tiada maklumat agen."
'End If
If Frm102.L34_Text.Visible = False Then
    If Frm102.TB11 = vbNullString Or (Frm102.TB11 <> vbNullString And Not IsNumeric(Frm102.TB11)) Then
        x = x + 1
        Err(x) = "Sila masukkan [Harga Semasa Emas 999.9]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm102.TB11 <> vbNullString And IsNumeric(Frm102.TB11) Then
    
        If Format(Frm102.TB11, "0.00") = "0.00" Then
            x = x + 1
            Err(x) = "Harga semasa emas 999.9 yang tidak sah. Nilai 0.00 tidak dibenarkan."
        End If
        
    End If

    If Frm102.CB5 = 0 And Frm102.CB6 = 0 Then
        x = x + 1
        Err(x) = "Sila buat pilihan cukai GST"
    End If
End If

If x <> 0 Then
    Frm6.Show
    Frm6.Pic1.Cls
    For Y = 1 To x
        Frm6.Pic1.Print Y & " - " & Err(Y)
    Next Y
    Exit Sub
Else

    If Frm102.L46_Text = vbNullString Then
    
        Note = "TIADA maklumat bagi agen yang diisi." & vbCrLf & _
                vbNullString & vbCrLf & _
                "Maklumat agen tidak akan dicetak di dalam invoice." & vbCrLf & _
                vbNullString & vbCrLf & _
                "Adakah anda yakin untuk teruskan urusan jualan ini ?"
    
    Else


        Note = "Adakah anda yakin untuk teruskan urusan jualan ini ?" & vbCrLf & _
                vbNullString & vbCrLf & _
                "Data jualan akan disimpan ke dalam sistem."
                
    End If
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbNo Then
        Exit Sub
    End If
    If Answer = vbYes Then
    
' ### Periksa kategori pembeli ### - Start
        If Frm102.L46_Text <> vbNullString Then
            If Frm28.L5_Text <> vbNullString Then
                
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from senarai_pelanggan where no_pelanggan='" & Frm28.L5_Text & "'", cn, adOpenKeyset, adLockOptimistic
            
                If Not rs.EOF Then
                
                    If Not IsNull(rs!kategori_pelanggan) Then Frm102_LM_KATEGORI = rs!kategori_pelanggan
                    
                End If
                
                rs.Close
                Set rs = Nothing
                
            End If
        End If
' ### Periksa kategori pembeli ### - End

        Frm102_LM_No_RESIT_JUALAN = Frm102.L23_Text 'No. Invoice
        Frm102_LM_No_VOUCHER = Frm102.L22_Text 'No. Voucher
        
Re_Gen_No_Rujukan:
    '###Masukkan Jualan Ke Dalam Table Akaun Jualan### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 22_jualan where no_resit='" & "BK" & Format(Frm102_LM_No_RESIT_JUALAN, "000000") & "'", cn, adOpenKeyset, adLockOptimistic
        
        If rs.EOF Then
            rs.AddNew
            If Frm102.L23_Text <> vbNullString Then
                rs!no_resit = "BK" & Format(Frm102_LM_No_RESIT_JUALAN, "000000") 'No. Resit Jualan
            Else
                rs!no_resit = Null 'No. Resit Jualan
            End If
            If Frm102.DTPicker1 <> vbNullString Then
                rs!tarikh = Frm102.DTPicker1 'Tarikh Jualan
            Else
                rs!tarikh = Null 'Tarikh Jualan
            End If
            
            rs!tunai = Format(0, "0.00")
            rs!bank_in = Format(0, "0.00") 'Cara Bayaran : Bank In
            rs!kad_kredit = Format(0, "0.00") 'Cara Bayaran : Kad Kredit
            rs!duit_simpanan_kedai = Format(0, "0.00") 'Cara Bayaran : Simpanan Duit Di Kedai
            rs!cek = Format(0, "0.00") 'Cara Bayaran : Cek
            
            '### Maklumat cara bayaran ### - Start
            If Frm102.TB14 <> vbNullString Then
                rs!tunai = Format(Frm102.TB14, "0.00") 'Cara Bayaran : Tunai
            Else
                rs!tunai = Null 'Cara Bayaran : Tunai
            End If
            If Frm102.TB15 <> vbNullString Then
                rs!bank_in = Format(Frm102.TB15, "0.00") 'Cara Bayaran : Bank In
            Else
                rs!bank_in = Null 'Cara Bayaran : Bank In
            End If
            If Frm102.TB16 <> vbNullString Then
                rs!kad_kredit = Format(Frm102.TB16, "0.00") 'Cara Bayaran : Kad Kredit
            Else
                rs!kad_kredit = Null 'Cara Bayaran : Kad Kredit
            End If
            If Frm102.L26_Text <> vbNullString Then
                rs!cas_Kad_Kredit = Frm102.L26_Text 'Cara Bayaran : Cas Kad Kredit (%)
            Else
                rs!cas_Kad_Kredit = Null 'Cara Bayaran : Cas Kad Kredit (%)
            End If
            If Frm102.TB17 <> vbNullString Then
                rs!jumlah_cas_kad_kredit = Format(Frm102.TB17, "0.00") 'Cara Bayaran : Jumlah Cas Kad Kredit (RM)
            Else
                rs!jumlah_cas_kad_kredit = Null 'Cara Bayaran : Jumlah Cas Kad Kredit (RM)
            End If
            If Frm102.TB18 <> vbNullString Then
                rs!jumlah_potongan_kad_kredit = Format(Frm102.TB18, "0.00") 'Cara Bayaran : Jumlah Potongan Kad Kredit (RM)
            Else
                rs!jumlah_potongan_kad_kredit = Null 'Cara Bayaran : Jumlah Potongan Kad Kredit (RM)
            End If
            If Frm102.TB22 <> vbNullString Then
                If Format(Frm102.TB22, "0.00") <> "0.00" Then
                    Frm102_LM_Flag_SIMPANAN = 1 '0 : Tiada Penggunakan Duit Simpanan Kedai , 1 : Ada Penggunakan Duit Simpanan Kedai
                End If
                rs!duit_simpanan_kedai = Format(Frm102.TB22, "0.00") 'Cara Bayaran : Simpanan Duit Di Kedai
            Else
                rs!duit_simpanan_kedai = Null 'Cara Bayaran : Simpanan Duit Di Kedai
            End If
            If Frm102.TB19 <> vbNullString Then
                rs!kad_debit = Format(Frm102.TB19, "0.00") 'Cara Bayaran : Kad Debit
            Else
                rs!kad_debit = Null 'Cara Bayaran : Kad Debit
            End If
            If Frm102.L27_Text <> vbNullString Then
                rs!cas_kad_debit = Frm102.L27_Text 'Cara Bayaran : Jumlah Cas Kad Debit (%)
            Else
                rs!cas_kad_debit = Null 'Cara Bayaran : Jumlah Cas Kad Debit (%)
            End If
            If Frm102.TB20 <> vbNullString Then
                rs!jumlah_cas_kad_debit = Format(Frm102.TB20, "0.00") 'Cara Bayaran : Jumlah Cas Kad Debit (RM)
            Else
                rs!jumlah_cas_kad_debit = Null 'Cara Bayaran : Jumlah Cas Kad Debit (RM)
            End If
            If Frm102.TB21 <> vbNullString Then
                rs!jumlah_potongan_kad_debit = Format(Frm102.TB21, "0.00") 'Cara Bayaran : Jumah Potongan Kad Debit (RM)
            Else
                rs!jumlah_potongan_kad_debit = Null 'Cara Bayaran : Jumah Potongan Kad Debit (RM)
            End If
            If Frm102.TB23 <> vbNullString Then
                rs!jumlah_bayaran = Format(Frm102.TB23, "0.00") 'Cara Bayaran : Jumlah Bayaran
            Else
                rs!jumlah_bayaran = Null 'Cara Bayaran : Jumlah Bayaran
            End If
            If Frm102.L15_Text <> vbNullString Then
                rs!harga_barang = Format(Frm102.L15_Text, "0.00") 'Jumlah Harga Barang Tanpa GST (RM)
            Else
                rs!harga_barang = Null 'Jumlah Harga Barang Tanpa GST (RM)
            End If
            If IsNumeric(Frm102.L19_Text) Then
                Frm102_LM_CUKAI_ZR = Frm102.L19_Text 'Maklumat GST : Jumlah GST ZR
            End If
            If IsNumeric(Frm102.L20_Text) Then
                Frm102_LM_CUKAI_SR = Frm102.L20_Text 'Maklumat GST : Jumlah GST SR
            End If
            rs!jumlah_cukai_gst = Format(Frm102_LM_CUKAI_ZR + Frm102_LM_CUKAI_SR, "0.00") 'Jumlah Cukai GST (ZR + SR)
            If Frm102.L16_Text <> vbNullString Then
                rs!harga_barang_dengan_gst = Format(Frm102.L16_Text, "0.00") 'Jumlah Harga Barang Dengan GST (RM)
                rs!harga_lepas_diskaun = Format(Frm102.L16_Text, "0.00") 'Harga Selepas Diskaun (RM)
                rs!harga_jualan = Format(Frm102.L16_Text, "0.00") 'Jumlah Harga Jualan (RM)
                rs!jumlah_perlu_bayar = Format(Frm102.L16_Text, "0.00") 'Jumlah Bayaran Yang Perlu Dibuat (RM)
            Else
                rs!harga_barang_dengan_gst = Null 'Jumlah Harga Barang Dengan GST (RM)
                rs!harga_lepas_diskaun = Null 'Harga Selepas Diskaun (RM)
                rs!harga_jualan = Null 'Jumlah Harga Jualan (RM)
                rs!jumlah_perlu_bayar = Null 'Jumlah Bayaran Yang Perlu Dibuat (RM)
            End If
            '### Maklumat cara bayaran ### - End
            
            rs!diskaun = Format(0, "0.00") 'Jumlah Diskaun (%)
            rs!adjustment = Format(0, "0.00") 'Adjustment (RM)
            rs!loss_trade_in = Format(0, "0.00") 'Potongan Harga Bagi Trade In Jika Kedai Perlu Bayar (%)
            rs!loss_trade_in_rm = Format(0, "0.00") 'Potongan Harga Bagi Trade In Jika Kedai Perlu Bayar (RM)
            rs!flag_bayaran = 0 '0 : Pembeli Bayar , 1 : Kedai Bayar
            rs!kuantiti_barang = Null 'Kuantiti Barang Yang Dijual
            rs!JUMLAH_BERAT = Null 'Jumlah Berat Barang Yang Dijual
            If Frm102.L17_Text <> vbNullString Then
                rs!gst_zr_harga = Format(Frm102.L17_Text, "0.00") 'Harga Keseluruhan Bagi Barang ZR
            Else
                rs!gst_zr_harga = Null 'Harga Keseluruhan Bagi Barang ZR
            End If
            If Frm102.L19_Text <> vbNullString Then
                rs!gst_zr_cukai = Format(Frm102.L19_Text, "0.00") 'Jumlah Cukai Bagi ZR
            Else
                rs!gst_zr_cukai = Null 'Jumlah Cukai Bagi ZR
            End If
            If Frm102.L18_Text <> vbNullString Then
                rs!gst_sr_harga = Format(Frm102.L18_Text, "0.00") 'Harga Keseluruhan Bagi Barang SR
            Else
                rs!gst_sr_harga = Null 'Harga Keseluruhan Bagi Barang SR
            End If
            If Frm102.L20_Text <> vbNullString Then
                rs!gst_sr_cukai = Format(Frm102.L20_Text, "0.00") 'Jumlah Cukai Bagi SR
            Else
                rs!gst_sr_cukai = Null 'Jumlah Cukai Bagi SR
            End If
            If Frm102.CBB4 <> vbNullString Then
                Frm102_LM_EMP_NO = Split(Frm102.CBB4, "  |  ")(1)
                rs!no_pekerja = Frm102_LM_EMP_NO 'No. Pekerja
            End If
            If Frm102.L46_Text <> vbNullString Then
                If Frm28.L5_Text <> vbNullString Then
                    rs!no_rujukan_pembeli = Frm28.L5_Text 'No. Rujukan Pembeli
                Else
                    rs!no_rujukan_pembeli = Null 'No. Rujukan Pembeli
                End If
            Else
                rs!no_rujukan_pembeli = Null 'No. Rujukan Pembeli
            End If
            If Frm102.L43_Text <> vbNullString Then
                rs!kuantiti_barang = Frm102.L43_Text 'Kuantiti barang
            Else
                rs!kuantiti_barang = 0 'Kuantiti barang
            End If
            If Frm102.L48_Text <> vbNullString Then
                rs!JUMLAH_BERAT = Frm102.L48_Text 'Jumlah berat
            Else
                rs!JUMLAH_BERAT = 0 'Kuantiti barang
            End If
            rs!no_rujukan_agen_dropship = Null 'No. Rujukan Agen Dropship
            rs!flag_trade_in = 0 '0 : Tiada Urusan Trade in , 1 : Ada Urusan Trade In
            rs!no_resit_trade_in = Null 'No. Resit Trade In
            rs!jumlah_trade_in = Null 'Jumlah Resit Trade In (RM)
            rs!jualan_online = 0
'1:  Pelanggan
'2:  Member
'3:  RAF
'4:  Pengedar
'5:  Normal Dealer
'6:  Master Dealer
    

            rs!kategori_pembeli = Frm102_LM_KATEGORI

            rs!invoice_type = 0 '0 : Unlimited , Selain 0 (Limited : Mengikut nombor yang dimasukkan)
            rs!epp = 0
            rs!approval_code_epp = Null
            rs!write_timestamp = Now
            DATA_SAVE = 1
            
            rs.Update
        Else
            Frm102_LM_No_RESIT_JUALAN = Frm102_LM_No_RESIT_JUALAN + 1
            Frm102.L23_Text = Frm102_LM_No_RESIT_JUALAN 'No. Invoice Jualan
            
            rs.Close
            Set rs = Nothing
            GoTo Re_Gen_No_Rujukan:
        End If
        
        rs.Close
        Set rs = Nothing
    '###Masukkan Jualan Ke Dalam Table Akaun Jualan### - End
    
'###Update Data Simpanan Duit Pelanggan### - Start
        'If Frm102_LM_Flag_SIMPANAN = 1 Then  '0 : Tiada Penggunakan Duit Simpanan Kedai , 1 : Ada Penggunakan Duit Simpanan Kedai
        
        '    Set rs = New ADODB.Recordset
        '    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        '    rs.Open "select * from senarai_pelanggan where no_pelanggan='" & Frm28.L5_Text & "'", cn, adOpenKeyset, adLockOptimistic
            
        '    If Not rs.EOF Then
        '        Frm102_LM_JUMLAH_SIMPANAN = Frm102.L26_Text  'Jumlah Simpanan Yang Ada
        '        Frm102_LM_GUNA_SIMPAN = Frm102.TB21  'Jumlah Simpanan Yang Hendak Digunakan
                
        '        rs!baki_simpanan = Format(Frm102_LM_JUMLAH_SIMPANAN - Frm102_LM_GUNA_SIMPAN, "0.00") 'Baki Simpanan
        '        rs.Update
        '    End If
            
        '    rs.Close
        '    Set rs = Nothing
            
        '    Set rs = New ADODB.Recordset
        '    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        '    rs.Open "select * from 24_rekod_kewangan_pelanggan", cn, adOpenKeyset, adLockOptimistic
            
        '    rs.AddNew
        '    rs!tarikh = Frm102.DTPicker1 'Tarikh
        '    rs!jenis = 1 '0 : Simpanan , 1 : Penggunaan Duit
        '    rs!no_rujukan_pelanggan = Frm28.L5_Text 'No. Rujukan Pelanggan
        '    rs!no_resit = "BK" & Format(Frm102_LM_No_RESIT_JUALAN, "000000") 'No. Resit Jualan
        '    rs!jumlah = Format(Frm102.TB22, "0.00") 'Jumlah Simpanan Yang Digunakan (RM)
        '    rs!jenis_penggunaan = 0 '0 : Belian Barangan Kemas , 1 : Ansuran , 2 : Tempahan (Deposit) , 3 : Servis , 4 : Tempahan (Ambilan Barang)
        '    rs.Update
            
        '    rs.Close
        '    Set rs = Nothing
           
        'End If
'###Update Data Simpanan Duit Pelanggan### - End

'### Masukkan Data Jualan Ke Dalam Table Jualan ### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from " & G_JUALAN_TEMP & " where status='" & "1" & "'", cn, adOpenKeyset, adLockOptimistic
        
        While rs.EOF = False
            Set rs1 = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs1.Open "select * from 23_senarai_jualan", cn, adOpenKeyset, adLockOptimistic
        
            rs1.AddNew
            rs1!tarikh = Frm102.DTPicker1 'Tarikh Jualan
            rs1!no_resit = "BK" & Format(Frm102_LM_No_RESIT_JUALAN, "000000") 'No. Invoice Jualan
            If Not IsNull(rs!no_siri_Produk) Then
                rs1!no_siri_Produk = rs!no_siri_Produk 'No. Siri Produk
            Else
                rs1!no_siri_Produk = Null 'No. Siri Produk
            End If
            If Not IsNull(rs!kategori_Produk) Then
                rs1!kategori_Produk = rs!kategori_Produk 'Kategori Produk
            Else
                rs1!no_siri_Produk = Null 'Kategori Produk
            End If
            If Not IsNull(rs!purity) Then
                rs1!purity = rs!purity 'Purity
            Else
                rs1!purity = Null 'Purity
            End If
            If Not IsNull(rs!Berat_Asal) Then
                rs1!Berat_Asal = rs!Berat_Asal 'Berat Asal (g)
            Else
                rs1!Berat_Asal = Null 'Berat Asal (g)
            End If
            If Not IsNull(rs!berat_jualan) Then
                rs1!berat_jualan = Format(rs!berat_jualan, "0.00") 'Berat Jualan (g)
            Else
                rs1!berat_jualan = Null 'Berat Jualan (g)
            End If
            If Not IsNull(rs!harga_Semasa) Then
                rs1!harga_Semasa = Format(rs!harga_Semasa, "0.00") 'Harga Semasa (RM/g)
            Else
                rs1!harga_Semasa = Null 'Harga Semasa (RM/g)
            End If
            If Not IsNull(rs!UPAH) Then
                rs1!UPAH = Format(rs!UPAH, "0.00") 'Upah (RM)
            Else
                rs1!UPAH = Null 'Upah (RM)
            End If
            If Not IsNull(rs!harga_asal) Then
                rs1!harga_asal = Format(rs!harga_asal, "0.00") 'Harga Asal Item (RM)
            Else
                rs1!harga_asal = Null 'Harga Asal Item (RM)
            End If
            If Not IsNull(rs!diskaun) Then
                rs1!diskaun = Format(rs!diskaun, "0.00") 'Diskaun (%)
            Else
                rs1!diskaun = Null 'Diskaun (%)
            End If
            If Not IsNull(rs!harga_lepas_diskaun) Then
                rs1!harga_lepas_diskaun = Format(rs!harga_lepas_diskaun, "0.00") 'Harga Selepas Diskaun (RM)
            Else
                rs1!harga_lepas_diskaun = Null 'Harga Selepas Diskaun (RM)
            End If
            If Not IsNull(rs!adjustment) Then
                rs1!adjustment = Format(rs!adjustment, "0.00") 'Harga Selepas Diskaun (RM)
            Else
                rs1!adjustment = Null 'Harga Selepas Diskaun (RM)
            End If
            If Not IsNull(rs!harga_jualan) Then
                rs1!harga_jualan = Format(rs!harga_jualan, "0.00") 'Harga Jualan (RM)
            Else
                rs1!harga_jualan = Null 'Harga Jualan (RM)
            End If
            If Not IsNull(rs!gst_ari_nashi) Then
                rs1!gst_ari_nashi = rs!gst_ari_nashi '0 : Tiada GST : ZR(L) , 1 : Ada GST : SR
            Else
                rs1!gst_ari_nashi = Null '0 : Tiada GST : ZR(L) , 1 : Ada GST : SR
            End If
            If Not IsNull(rs!kadar_gst) Then
                rs1!kadar_gst = rs!kadar_gst 'Kadar Cukai GST (%)
            Else
                rs1!kadar_gst = Null 'Kadar Cukai GST (%)
            End If
            If Not IsNull(rs!jumlah_gst) Then
                rs1!jumlah_gst = Format(rs!jumlah_gst, "0.00") 'Jumlah Cukai GST (RM)
            Else
                rs1!jumlah_gst = Null 'Jumlah Cukai GST (RM)
            End If
            If Not IsNull(rs!harga_dengan_gst) Then
                rs1!harga_dengan_gst = Format(rs!harga_dengan_gst, "0.00") 'Harga Jualan Termasuk GST (RM)
            Else
                rs1!harga_dengan_gst = Null 'Harga Jualan Termasuk GST (RM)
            End If
            If Not IsNull(rs!dropship) Then
                rs1!dropship = rs!dropship '0 : Jualan Bukan Oleh Agen Dropship , 1 : Jualan Oleh Agen Dropship
            Else
                rs1!dropship = Null '0 : Jualan Bukan Oleh Agen Dropship , 1 : Jualan Oleh Agen Dropship
            End If
            If Not IsNull(rs!komisyen_per_gram) Then
                rs1!komisyen_per_gram = Format(rs!komisyen_per_gram, "0.00") 'Komisyen Per Gram Dropship (RM/g) : BK
            Else
                rs1!komisyen_per_gram = Null 'Komisyen Per Gram Dropship (RM/g) : BK
            End If
            If Not IsNull(rs!jumlah_komisyen) Then
                rs1!jumlah_komisyen = Format(rs!jumlah_komisyen, "0.00") 'Jumlah Komisyen Kepada Agen Dropship Bagi Item Ini : BK
            Else
                rs1!jumlah_komisyen = Null 'Jumlah Komisyen Kepada Agen Dropship Bagi Item Ini : BK
            End If
            If Not IsNull(rs!harga_per_gram_modal) Then
                rs1!harga_per_gram_modal = Format(rs!harga_per_gram_modal, "0.00") 'Harga Per Gram Bagi Modal (RM/g)
            Else
                rs1!harga_per_gram_modal = Null 'Harga Per Gram Bagi Modal (RM/g)
            End If
            If Not IsNull(rs!modal) Then
                rs1!modal = Format(rs!modal, "0.00") 'Harga Modal (RM)
            Else
                rs1!modal = Null 'Harga Modal (RM)
            End If
            If Not IsNull(rs!untung) Then
                rs1!untung = Format(rs!untung, "0.00") 'Jumlah Keuntungan
            Else
                rs1!untung = Null 'Jumlah Keuntungan
            End If
            If Not IsNull(rs!harga_per_gram_supplier) Then
                rs1!harga_per_gram_supplier = Format(rs!harga_per_gram_supplier, "0.00") 'Harga per gram (harga semasa) dari supplier (modal)
            Else
                rs1!harga_per_gram_supplier = Null 'Harga per gram (harga semasa) dari supplier (modal)
            End If
            If Not IsNull(rs!upah_modal) Then
                rs1!upah_modal = Format(rs!upah_modal, "0.00") 'Upah modal
            Else
                rs1!upah_modal = Null 'Upah modal
            End If
            If Not IsNull(rs!untung2) Then
                rs1!untung2 = Format(rs!untung2, "0.00") 'Jumlah Keuntungan
            Else
                rs1!untung2 = Null 'Jumlah Keuntungan
            End If
            If Not IsNull(rs!dulang) Then
                rs1!dulang = rs!dulang 'Dulang
            Else
                rs1!dulang = Null 'Dulang
            End If
            If Not IsNull(rs!potong_flag) Then
                rs1!potong_flag = rs!potong_flag '0 : Tiada Potong , 1 : Ada Potong
                If rs!potong_flag = 0 Then
                    rs1!Status = 0 '0 : Jualan Biasa , 1 : Jualan Secara Potong , 2 : Tempahan , 3 : Ansuran , 4 : ETA
                Else
                    rs1!Status = 1 '0 : Jualan Biasa , 1 : Jualan Secara Potong , 2 : Tempahan , 3 : Ansuran , 4 : ETA
                End If
            Else
                rs1!potong_flag = Null '0 : Tiada Potong , 1 : Ada Potong
            End If
            If Not IsNull(rs!Type) Then
                rs1!Type = rs!Type '0 : BK , 1 : Barang Permata
            Else
                rs1!Type = Null '0 : BK , 1 : Barang Permata
            End If
            If Frm102.CBB4 <> vbNullString Then
                Frm102_LM_EMP_NO = Split(Frm102.CBB4, "  |  ")(1)
                rs1!no_pekerja = Frm102_LM_EMP_NO 'No. Pekerja
            End If
            If Frm102.L46_Text <> vbNullString Then
                If Frm28.L5_Text <> vbNullString Then
                    rs1!no_rujukan_pembeli = Frm28.L5_Text 'No. Rujukan Pembeli
                Else
                    rs1!no_rujukan_pembeli = Null 'No. Rujukan Pembeli
                End If
            Else
                rs1!no_rujukan_pembeli = Null 'No. Rujukan Pembeli
            End If
            rs1!jualan_online = 0
            'If Frm27.L5_Text <> vbNullString Then
            '    rs1!no_rujukan_agen_dropship = Frm27.L5_Text 'No. Rujukan Agen Dropship
            'Else
            '    rs1!no_rujukan_agen_dropship = Null 'No. Rujukan Agen Dropship
            'End If
            
'1:  Pelanggan
'2:  Member
'3:  RAF
'4:  Pengedar
'5:  Normal Dealer
'6:  Master Dealer

            'If Frm102.CB4 = 1 Then
            '    rs1!kategori_pembeli = 1
            'ElseIf Frm102.CB5 = 1 Then
            '    rs1!kategori_pembeli = 2
            'ElseIf Frm102.CB6 = 1 Then
            '    rs1!kategori_pembeli = 4
            'ElseIf Frm102.CB9 = 1 Then
            rs1!kategori_pembeli = Frm102_LM_KATEGORI
            'ElseIf Frm102.CB10 = 1 Then
            '    rs1!kategori_pembeli = 5
            'ElseIf Frm102.CB11 = 1 Then
            '    rs1!kategori_pembeli = 6
            'End If
            
            If Not IsNull(rs!gst_include) Then 'Pilihan Cukai GST (SR) Samada Pelanggan Bayar Atau Kedai Bayar
                If rs!gst_include = 0 Then
                    rs1!gst_include = Null
                ElseIf rs!gst_include = 1 Then
                    rs1!gst_include = "**Harga Termasuk GST"
                End If
            Else
                rs1!gst_include = Null
            End If
            If Not IsNull(rs!harga_tanpa_gst) Then
                rs1!harga_tanpa_gst = Format(rs!harga_tanpa_gst, "0.00") 'Harga Semasa (RM/g)
            Else
                rs1!harga_tanpa_gst = Null 'Harga Semasa (RM/g)
            End If

'### Maklumat tetapan harga jualan kepada staff ### - Start
            If Not IsNull(rs!kadar_penurunan_upah) Then 'Kadar peratusan penurunan harga upah kepada staff (%)
                rs1!kadar_penurunan_upah = Format(rs!kadar_penurunan_upah, "0.00")
            Else
                rs1!kadar_penurunan_upah = Null
            End If
            If Not IsNull(rs!harga_semasa_staff) Then 'Harga emas semasa yang dijual kepada staff
                rs1!harga_semasa_staff = Format(rs!harga_semasa_staff, "0.00")
            Else
                rs1!harga_semasa_staff = Null
            End If
            If Not IsNull(rs!kadar_penurunan_bp) Then 'Kadar peratusan penurunan harga barang permata kepada staff (%)
                rs1!kadar_penurunan_bp = Format(rs!kadar_penurunan_bp, "0.00")
            Else
                rs1!kadar_penurunan_bp = Null
            End If
            If Not IsNull(rs!harga_staff) Then 'Harga yang dijual kepada staff (RM)
                rs1!harga_staff = Format(rs!harga_staff, "0.00")
            Else
                rs1!harga_staff = Null
            End If
            If Not IsNull(rs!harga_bp_asal) Then 'Tetapan harga barang permata yang asal (RM)
                rs1!harga_bp_asal = Format(rs!harga_bp_asal, "0.00")
            Else
                rs1!harga_bp_asal = Null
            End If
            If Not IsNull(rs!upah_asal) Then 'Tetapan upah asal (RM)
                rs1!upah_asal = Format(rs!upah_asal, "0.00")
            Else
                rs1!upah_asal = Null
            End If
            If Not IsNull(rs!komisyen_staff) Then 'Tetapan upah asal (RM)
                rs1!komisyen_staff = Format(rs!komisyen_staff, "0.00")
            Else
                rs1!komisyen_staff = Null
            End If
'### Maklumat tetapan harga jualan kepada staff ### - End

            If Not IsNull(rs!pemalar_tukaran_999) Then 'Pemalar tukaran emas kepada 999.9 bagi urusan jualan
                rs1!pemalar_tukaran_999 = rs!pemalar_tukaran_999
            Else
                rs1!pemalar_tukaran_999 = Null
            End If
            If Not IsNull(rs!berat_999) Then 'Berat jualan dalam purity 999.9
                rs1!berat_999 = Format(rs!berat_999, "0.00")
            Else
                rs1!berat_999 = Null
            End If
            rs1!jenis_jualan = 1 '0 : Jualan biasa kepada pelanggan , 1 : Jualan secara tukaran barang kepada agen
            If Not IsNull(rs!gst_barang_atau_upah) Then '0 : GST pada harga jualan , 1 : GST pada upah
                rs1!gst_barang_atau_upah = rs!gst_barang_atau_upah
            Else
                rs1!gst_barang_atau_upah = 0
            End If
            If Not IsNull(rs!harga_jualan_dengan_gst) Then
                rs1!harga_jualan_dengan_gst = rs!harga_jualan_dengan_gst
            Else
                rs1!harga_jualan_dengan_gst = 0
            End If
            rs1!write_timestamp = Now
            
            rs1.Update
            
            rs1.Close
            Set rs1 = Nothing
            
'### Update Table Database Bagi Item Ini ### - Start
            Set rs2 = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs2.Open "select * from Data_Database where no_siri_produk='" & rs!no_siri_Produk & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs2.EOF Then
                If rs!Type = 0 Then
                    Frm102_LM_BERAT_ASAL = rs2!beza_berat 'Berat Asal (g)
                    Frm102_LM_BERAT_JUALAN = rs!berat_jualan 'Berat Jualan (g)
                    
                    If Frm102_LM_BERAT_JUALAN = Frm102_LM_BERAT_ASAL Then
                        rs2!beza_berat = "0.00" 'Baki Berat
                        rs2!StatusItem = 11
                    Else
                        rs2!beza_berat = Format(Frm102_LM_BERAT_ASAL - Frm102_LM_BERAT_JUALAN, "0.00") 'Baki Berat
                        rs2!StatusItem = 12
                    End If
                Else
                    rs2!StatusItem = 11
                End If
                rs2.Update
            End If
            
            rs2.Close
            Set rs2 = Nothing
'### Update Table Database Bagi Item Ini ### - End

            rs.MoveNext
        Wend
        
        rs.Close
        Set rs = Nothing
        
'### Masukkan data belian barang dari agen ke dalam database ### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 49_belian_temp where status='" & "1" & "'", cn, adOpenKeyset, adLockOptimistic
        
        While rs.EOF = False
            Set rs1 = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs1.Open "select * from 50_belian_emas_agen", cn, adOpenKeyset, adLockOptimistic
        
            rs1.AddNew
            rs1!no_invoice = "BK" & Format(Frm102_LM_No_RESIT_JUALAN, "000000") 'No. Invoice
            rs1!tarikh = Frm102.DTPicker1 'Tarikh Jualan
            If Not IsNull(rs!Berat_Asal) Then rs1!Berat_Asal = rs!Berat_Asal 'Berat asal barang
            If Not IsNull(rs!purity) Then rs1!purity = rs!purity 'Purity barang
            If Not IsNull(rs!kod_Purity) Then rs1!kod_Purity = rs!kod_Purity 'Kod purity barang
            If Not IsNull(rs!kadar_tukaran) Then rs1!kadar_tukaran = rs!kadar_tukaran 'Kadar tukaran kepada purity 999.9
            If Not IsNull(rs!berat_tukaran) Then rs1!berat_tukaran = rs!berat_tukaran 'Berat setelah ditukar kepada purity 999.9
            If Not IsNull(rs!Status) Then rs1!Status = rs!Status 'Status
            rs1!write_timestamp = Now
            
            rs1.Update
            
            rs1.Close
            Set rs1 = Nothing
            
            rs.MoveNext
        Wend
        
        rs.Close
        Set rs = Nothing
'### Masukkan data belian barang dari agen ke dalam database ### - End
        
'### Periksa samada ada pembayaran kepada agen atau tidak bagi urusan belian barang kemas ### - Start
        If Frm102.L34_Text.Visible = True Then
Re_Gen_No_Rujukan2:
'###Masukkan Jualan Ke Dalam Table Akaun Jualan### - Start
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from 51_voucher_belian_agen where no_voucher='" & "TIA" & Format(Frm102_LM_No_VOUCHER, "000000") & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                Frm102_LM_No_VOUCHER = Frm102_LM_No_VOUCHER + 1
                Frm102.L22_Text = Frm102_LM_No_VOUCHER 'No. Invoice Jualan
                
                rs.Close
                Set rs = Nothing
                
                GoTo Re_Gen_No_Rujukan2:
            End If
            
            rs.Close
            Set rs = Nothing
        End If
'### Periksa samada ada pembayaran kepada agen atau tidak bagi urusan belian barang kemas ### - End

'### Masukkan data voucher / invoice bagi belian agen ini ### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 51_voucher_belian_agen", cn, adOpenKeyset, adLockOptimistic
        
        'If rs.EOF Then
            rs.AddNew
            If Frm102.L34_Text.Visible = True Then 'No. Voucher
                rs!no_voucher = "TIA" & Format(Frm102_LM_No_VOUCHER, "000000")
                rs!flag_bayaran = 1 '0 : Bayaran dibuat oleh pembeli , 1 : Bayaran dibuat oleh pihak kedai
            Else
                rs!no_voucher = Null
                rs!flag_bayaran = 0 '0 : Bayaran dibuat oleh pembeli , 1 : Bayaran dibuat oleh pihak kedai
            End If
            rs!no_invoice = "BK" & Format(Frm102_LM_No_RESIT_JUALAN, "000000") 'No. Invoice
            rs!tarikh = Frm102.DTPicker1 'Tarikh belian
            If Frm102.L9_Text <> vbNullString Then 'Berat jualan keseluruhan barang kedai
                rs!berat_jualan = Format(Frm102.L9_Text, "0.00")
            Else
                rs!berat_jualan = "0.00"
            End If
            If Frm102.L10_Text <> vbNullString Then 'Berat belian keseluruhan (Barang trade in)
                rs!berat_belian = Format(Frm102.L10_Text, "0.00")
            Else
                rs!berat_belian = "0.00"
            End If
            If Frm102.L11_Text <> vbNullString Then 'Beza antara berat jualan dan belian
                rs!beza_berat = Format(Frm102.L11_Text, "0.00")
            Else
                rs!beza_berat = "0.00"
            End If
            If Frm102.TB11 <> vbNullString Then 'Harga semasa (penilaian harga emas oleh pihak kedai)
                rs!harga_Semasa = Format(Frm102.TB11, "0.00")
            Else
                rs!harga_Semasa = "0.00"
            End If
            If Frm102.L12_Text <> vbNullString Then 'Nilaian harga emas oleh pihak kedai terhadap beza berat tersebut (jika bayaran perlu dibuat oleh pihak kedai sahaja)
                rs!harga_emas = Format(Frm102.L12_Text, "0.00")
            Else
                rs!harga_emas = "0.00"
            End If
            If Frm102.L31_Text <> vbNullString Then 'Harga emas tanpa GST
                rs!harga_tanpa_gst = Format(Frm102.L31_Text, "0.00")
            Else
                rs!harga_tanpa_gst = "0.00"
            End If
            
            If Frm102.CB5 = 1 Then
                rs!gst_ari_nashi = "ZR (L)" '0 : Tiada GST : ZR(L) , 1 : Ada GST : SR
                rs!gst_include = 0 '0 : GST Dibayar Oleh Pelanggan , 1 : GST Termasuk Dalam Harga Barang
            ElseIf Frm102.CB6 = 1 Then
                rs!gst_ari_nashi = "SR" '0 : Tiada GST : ZR(L) , 1 : Ada GST : SR
                
                If Frm102.CB7 = 1 Then 'Jenis Cukai GST SR
                    rs!gst_include = 1 '0 : GST Dibayar Oleh Pelanggan , 1 : GST Termasuk Dalam Harga Barang
                Else
                    rs!gst_include = 0 '0 : GST Dibayar Oleh Pelanggan , 1 : GST Termasuk Dalam Harga Barang
                End If 'here2
            End If
            If Frm102.L21_Text <> vbNullString Then
                rs!kadar_gst = Frm102.L21_Text 'Kadar Cukai GST (%)
            Else
                rs!kadar_gst = "0" 'Jumlah Cukai GST (RM)
            End If
            If Frm102.L31_Text <> vbNullString Then 'Harga emas tanpa GST
                rs!harga_tanpa_gst = Format(Frm102.L31_Text, "0.00")
            Else
                rs!harga_tanpa_gst = "0.00"
            End If
            If Frm102.TB12 <> vbNullString Then 'Jumlah Cukai GST (RM)
                rs!jumlah_gst = Format(Frm102.TB12, "0.00")
            Else
                rs!jumlah_gst = "0.00"
            End If
            If Frm102.TB13 <> vbNullString Then 'Jumlah emas + GST (RM)
                rs!harga_dengan_gst = Format(Frm102.TB13, "0.00")
            Else
                rs!harga_dengan_gst = "0.00"
            End If
            If Frm102.CBB4 <> vbNullString Then
                Frm102_LM_EMP_NO = Split(Frm102.CBB4, "  |  ")(1)
                rs!no_pekerja = Frm102_LM_EMP_NO 'No. Pekerja
            End If
            rs!write_timestamp = Now
            
            rs.Update
        'End If
        
        rs.Close
        Set rs = Nothing
'### Masukkan data voucher / invoice bagi belian agen ini ### - End
    
        If DATA_SAVE = 1 Then
    '###Update No. Resit### - Start
            G_No_RESIT_JUALAN = vbNullString
            G_No_RESIT_JUALAN = "BK" & Format(Frm102_LM_No_RESIT_JUALAN, "000000")
            
    '#### Update Log Aktiviti Sistem #### - Start
            If Frm102.CBB4 <> vbNullString Then
                Frm102_LM_EMP_NAME = Split(Frm102.CBB4, "  |  ")(0)
            End If
        
            'User = MDI_frm1.L3_Text
            LogAct_Memory = "[" & Frm102_LM_EMP_NAME & "] Jualan Barang Kemas Kepada Agen. No. Invoice [" & "BK" & Format(Frm102_LM_No_RESIT_JUALAN, "000000") & "]."
            LogDate_Memory = DateTime.Date & " " & DateTime.Time$
            Call UpdateLog_Database
    '#### Update Log Aktiviti Sistem #### - End
    
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from default_setting", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                If rs!Default1 = "Default" Then
                    rs!ResitNo = Frm102_LM_No_RESIT_JUALAN + 1 'No. Invoice
                    If Frm102.L34_Text.Visible = True Then
                        rs!no_trade_in_agen = Frm102.L22_Text + 1 'No. Voucher Trade In
                    End If
                    rs.Update
                End If
            End If
            
            rs.Close
            Set rs = Nothing
            
            Call frm102_reset_main
            Call frm102_reset_1
            Call frm102_reset_2
            Call frm102_reset_3
            Unload Frm28
            Call Frm102_Senarai_Jualan_Header
            Call Frm102_Senarai_Jualan
            Call Frm102_senarai_belian_header
            Call Frm102_senarai_belian
            
            Frm102.TB1.SetFocus
            
            Call Frm102_cetak_invoice
            Call Frm102_cetak_voucher
    '###Update No. Resit### - End
        End If
        
    End If
End If
End Sub
Private Sub CMD9_Click()
'On Error Resume Next
Unload Frm102
MDI_frm1.L5_Text = 0
End Sub
Private Sub Form_Load()
'On Error Resume Next
Frm102.L43_Text = 0 'Jumlah bilangan barang jualan
Frm102.L44_Text = 0 'Jumlah bilangan barang trade in
Frm102.L48_Text = "0.00" 'Jumlah berat (g)
'GLOBAL_DISABLE = 0
'Frm102.TB1 = vbNullString

'Call frm102_reset_1
'Call frm102_reset_2
'Call frm102_reset_3
'Call frm102_reset_main

'Frm102.L26_Text.BackStyle = 0
'Frm102.L27_Text.BackStyle = 0

'Frm102.DTPicker1 = DateTime.Date$
End Sub
Private Sub Frm102_SM_edit_data1_Click()
'on error resume next
DATA_FOUND = 0

If Frm102.MSFlexGrid1 <> vbNullString Then
    Frm102_LM_ID = Frm102.MSFlexGrid1.TextMatrix(Frm102.MSFlexGrid1, 2) 'No. ID
    
    If Frm102_LM_ID <> vbNullString Then
        Call frm102_reset_1 '!! Hati-hati dengan tempat letakkan command ini!!
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from " & G_JUALAN_TEMP & " where ID='" & Frm102_LM_ID & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Not IsNull(rs!ID) Then Frm102.L24_Text = rs!ID 'No. ID Database
            If Not IsNull(rs!no_siri_Produk) Then Frm102.L3_Text = rs!no_siri_Produk 'No. Siri Produk
            If Not IsNull(rs!kategori_Produk) Then Frm102.L5_Text = rs!kategori_Produk 'Kategori Produk
            If Not IsNull(rs!purity) Then Frm102.L4_Text = rs!purity 'Purity
            If Not IsNull(rs!Berat_Asal) Then Frm102.L6_Text = Format(rs!Berat_Asal, "#,##0.00") 'Berat Asal (g)
            If Not IsNull(rs!berat_jualan) Then Frm102.TB3 = Format(rs!berat_jualan, "#,##0.00") 'Berat Jualan (g)
            If Not IsNull(rs!harga_Semasa) Then Frm102.TB2 = Format(rs!harga_Semasa, "#,##0.00") 'Harga Semasa (RM/g)
            If Not IsNull(rs!UPAH) Then Frm102.TB4 = Format(rs!UPAH, "#,##0.00") 'Upah (RM)
            If Not IsNull(rs!gst_ari_nashi) Then 'Harga Jualan (RM)
                If rs!gst_ari_nashi = "ZR (L)" Then
                    Frm102.CB2 = 1 '0 : Tiada GST : ZR(L) , 1 : Ada GST : SR
                    If Not IsNull(rs!jumlah_gst) Then
                        Frm102.TB5 = Format(rs!jumlah_gst, "#,##0.00") 'Jumlah Cukai GST (RM)
                    Else
                        Frm102.TB5 = "0.00"
                    End If
                ElseIf rs!gst_ari_nashi = "SR" Then
                    Frm102.CB3 = 1 '0 : Tiada GST : ZR(L) , 1 : Ada GST : SR
                    If Not IsNull(rs!kadar_gst) Then
                        Frm102.L21_Text = rs!kadar_gst 'Kadar Cukai GST (%)
                    End If
                    If Not IsNull(rs!jumlah_gst) Then
                        Frm102.TB5 = rs!jumlah_gst 'Jumlah Cukai GST (RM)
                    Else
                        Frm102.TB5 = "0.00"
                    End If
                    If Not IsNull(rs!gst_include) Then
                        If rs!gst_include = 0 Then
                            Frm102.CB4 = 0
                        ElseIf rs!gst_include = 1 Then
                            Frm102.CB4 = 1
                        End If
                    Else
                        Frm102.CB4 = 0
                    End If
                End If
            End If
            If Not IsNull(rs!harga_tanpa_gst) Then Frm102.L30_Text = Format(rs!harga_tanpa_gst, "#,##0.00") 'Harga Jualan Tanpa GST (RM)
            If Not IsNull(rs!harga_dengan_gst) Then Frm102.TB6 = Format(rs!harga_dengan_gst, "#,##0.00") 'Harga Jualan Termasuk GST (RM)
            If Not IsNull(rs!harga_per_gram_modal) Then Frm102.L33_Text = Format(rs!harga_per_gram_modal, "#,##0.00") 'Harga Per Gram Bagi Modal (RM/g)
            If Not IsNull(rs!berat_999) Then Frm102.L7_Text = rs!berat_999 'Berat jualan dalam purity 999.9
            If Not IsNull(rs!harga_per_gram_supplier) Then Frm102.L49_Text = Format(rs!harga_per_gram_supplier, "#,##0.00") 'Harga per gram (harga semasa) dari supplier (modal)
            If Not IsNull(rs!upah_modal) Then Frm102.L50_Text = Format(rs!upah_modal, "#,##0.00") 'Upah modal
            
            On Error GoTo Err_A:
            If Not IsNull(rs!pemalar_tukaran_999) Then
                Frm102_LM_PURITY = rs!pemalar_tukaran_999 'Purity
                Frm102.CBB1 = Frm102_LM_PURITY 'Purity
            End If
            
Restore_A:
            DATA_FOUND = 1
        End If
        
        rs.Close
        Set rs = Nothing
        
        If DATA_FOUND = 1 Then
            Frm102.CMD1.Visible = False
            Frm102.CMD2.Visible = True
            Frm102.CMD3.Visible = True
        End If
        
    End If
End If

Exit Sub
Err_A:
Frm102.CBB1.AddItem Frm102_LM_PURITY
Frm102.CBB1 = Frm102_LM_PURITY
Resume Restore_A:
End Sub
Private Sub Frm102_SM_edit_data2_Click()
'on error resume next
DATA_FOUND = 0

If Frm102.MSFlexGrid2 <> vbNullString Then
    Frm102_LM_ID = Frm102.MSFlexGrid2.TextMatrix(Frm102.MSFlexGrid2, 2) 'No. ID
    
    If Frm102_LM_ID <> vbNullString Then
        Call frm102_reset_2 '!! Hati-hati dengan tempat letakkan command ini!!
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 49_belian_temp where ID='" & Frm102_LM_ID & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            GLOBAL_DISABLE = 1
            
            If Not IsNull(rs!ID) Then Frm102.L25_Text = rs!ID 'No. ID Database
            If Not IsNull(rs!Berat_Asal) Then Frm102.TB10 = Format(rs!Berat_Asal, "#,##0.00") 'Berat (barang trade in)
            If Not IsNull(rs!kod_Purity) Then Frm102.L47_Text = rs!kod_Purity 'Kod purity
            
            On Error GoTo Err_A:
            If Not IsNull(rs!purity) Then
                Frm102_LM_PURITY = rs!purity 'Purity
                Frm102.CBB2 = Frm102_LM_PURITY 'Purity
            End If
            
Restore_A:
    
            'on error resume next

            If Not IsNull(rs!kadar_tukaran) Then Frm102.TB24 = rs!kadar_tukaran 'Trade In : Kadar tukaran purity 999.9
            If Not IsNull(rs!berat_tukaran) Then Frm102.L8_Text = Format(rs!berat_tukaran, "#,##0.00") 'Trade In : Berat dalam 999.9
            
            GLOBAL_DISABLE = 0
            DATA_FOUND = 1
        End If
        
        rs.Close
        Set rs = Nothing
        
        If DATA_FOUND = 1 Then

            Frm102.CMD4.Visible = False 'Masukkan dalam senarai trade in
            Frm102.CMD5.Visible = True 'Masukkan dalam senarai trade in (Edit)
            Frm102.CMD6.Visible = True 'Batal edit data

        End If
        
    End If
End If

Exit Sub
Err_A:
Frm102.CBB2.AddItem Frm102_LM_PURITY
Frm102.CBB2 = Frm102_LM_PURITY
Resume Restore_A:
End Sub
Private Sub Frm102_SM_remove_belian_Click()
'on error resume next
DATA_FOUND = 0

If Frm102.MSFlexGrid2 <> vbNullString Then
    Frm102_LM_ID = Frm102.MSFlexGrid2.TextMatrix(Frm102.MSFlexGrid2, 2) 'No. ID
    
    If Frm102_LM_ID <> vbNullString Then
        
        Note = "Adakah anda ingin keluarkan item ini dari senarai belian ?"
                
        Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
        If Answer = vbNo Then
            'Exit Sub
        End If
        If Answer = vbYes Then
        
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from 49_belian_temp where ID='" & Frm102_LM_ID & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                If Not IsNull(rs!Status) Then
'### Kod Bagi Status ###
'==========================================
'1 : Kemasukkan data baru (Data baru)
'2 : Tiada perubahan (Menu edit)
'3 : Data telah diedit (Menu edit)
'4 : Kemasukkan data baru (Menu edit)
'5 : Data dipadamkan (Data yang diterima dari menu baru)

                    If rs!Status = 1 Or rs!Status = 4 Then
                        rs.Delete
                        rs.Update
                        
                        DATA_FOUND = 1
                    ElseIf rs!Status = 2 Or rs!Status = 3 Then
                        rs!Status = 5
                        rs.Update
                        
                        DATA_FOUND = 1
                    End If
                    
                End If
            End If
            
            rs.Close
            Set rs = Nothing
            
            If DATA_FOUND = 1 Then
                Call frm102_reset_2
                
                Call Frm102_senarai_belian_header
                Call Frm102_senarai_belian
                
                MsgBox "Item telah dikeluarkan dari senarai jualan.", vbInformation, "Info"
            End If
        End If
        
    End If
End If
End Sub
Private Sub Frm102_SM_remove_jualan_Click()
'on error resume next
DATA_FOUND = 0

If Frm102.MSFlexGrid1 <> vbNullString Then
    Frm102_LM_ID = Frm102.MSFlexGrid1.TextMatrix(Frm102.MSFlexGrid1, 2) 'No. ID
    
    If Frm102_LM_ID <> vbNullString Then
        
        Note = "Adakah anda ingin keluarkan item ini dari senarai jualan dan pulangkan ke stok kedai ?"
                
        Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
        If Answer = vbNo Then
            'Exit Sub
        End If
        If Answer = vbYes Then
        
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from " & G_JUALAN_TEMP & " where ID='" & Frm102_LM_ID & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                If Not IsNull(rs!Status) Then
'### Kod Bagi Status ###
'==========================================
'0 : Tidak Aktif (Tidak Perlu Buat Apa-Apa)
'1 : Aktif - Kemasukkan Data Baru
'2 : Barang Sudah Terjual ****Tidak Dibenarkan Untuk Diedit Atau Dipadamkan
'3 : Kemasukkan Data Baru
'4 : Data Diedit
'5 : Keluarkan Data Dari Database (Bagi Item Asal Yang Dipadamkan)
'6 : Ignore Kemasukkan Data Ke Dalam Database
                    If rs!Status = 1 Or rs!Status = 4 Then
                        rs.Delete
                        rs.Update
                        
                        DATA_FOUND = 1
                    ElseIf rs!Status = 2 Or rs!Status = 3 Then
                        rs!Status = 5
                        rs.Update
                        
                        DATA_FOUND = 1
                    End If
                End If
            End If
            
            rs.Close
            Set rs = Nothing
            
            If DATA_FOUND = 1 Then
                Call frm102_reset_1
                
                Call Frm102_Senarai_Jualan_Header
                Call Frm102_Senarai_Jualan
                
                MsgBox "Item telah dikeluarkan dari senarai jualan.", vbInformation, "Info"
            End If
        End If
        
    End If
End If
End Sub
Private Sub L11_Text_Change()
'On Error Resume Next
Call frm102_calc5
End Sub
Private Sub L12_Text_Change()
'On Error Resume Next
Call frm102_calc8
End Sub
Private Sub L13_Text_Change()
'On Error Resume Next
Call frm102_calc7
End Sub
Private Sub L14_Text_Change()
'On Error Resume Next
If IsNumeric(Frm102.L14_Text) Then
    If Frm102.L34_Text.Visible = False Then
        Frm102.TB14 = Format(Frm102.L14_Text, "0.00")
    Else
        Frm102.TB14 = Format(0, "0.00")
    End If
Else
    Frm102.TB14 = Format(0, "0.00")
End If

Call frm102_calc10
End Sub
Private Sub L26_Text_Change()
'On Error Resume Next
Call frm102_calc12
End Sub
Private Sub L27_Text_Change()
'On Error Resume Next
Call frm102_calc14
End Sub
Private Sub L30_Text_Change()
'On Error Resume Next
Call frm102_calc3
End Sub
Private Sub L31_Text_Change()
'On Error Resume Next
Call frm102_calc9
End Sub

Private Sub L9_Text_Change()
'On Error Resume Next
Call frm102_calc4
End Sub
Private Sub L10_Text_Change()
'On Error Resume Next
Call frm102_calc4
End Sub

Private Sub Label21_Click()

End Sub

Private Sub MSFlexGrid1_DblClick()
'On Error Resume Next
If Frm102.MSFlexGrid1 <> vbNullString Then
    PopupMenu Frm102_PM_menu
Else
    MsgBox "Tiada Data.", vbExclamation, "Info"
End If
End Sub
Private Sub MSFlexGrid1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'On Error Resume Next
If Frm102.MSFlexGrid1 <> vbNullString Then
    If Button = vbRightButton Then
        PopupMenu Frm102_PM_menu, vbPopupMenuRightButton
    End If
End If
End Sub
Private Sub MSFlexGrid2_DblClick()
'On Error Resume Next
If Frm102.MSFlexGrid2 <> vbNullString Then
    PopupMenu Frm102_PM_menu2
Else
    MsgBox "Tiada Data.", vbExclamation, "Info"
End If
End Sub
Private Sub MSFlexGrid2_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'On Error Resume Next
If Frm102.MSFlexGrid2 <> vbNullString Then
    If Button = vbRightButton Then
        PopupMenu Frm102_PM_menu2, vbPopupMenuRightButton
    End If
End If
End Sub

Private Sub TB1_Change()
'on error resume next
If Frm102.CB1 = 1 And Frm102.TB1 <> vbNullString Then
    Frm102.Tmr2.Enabled = False
    Frm102.Tmr2.Enabled = True
    Frm102.Tmr2.Interval = 100
End If
End Sub
Private Sub TB10_Change()
'On Error Resume Next
Call frm102_calc6
End Sub
Private Sub TB11_Change()
'On Error Resume Next
Call frm102_calc5
End Sub
Private Sub TB12_Change()
'On Error Resume Next
Call frm102_calc9
End Sub
Private Sub TB13_Change()
'On Error Resume Next
Call frm102_calc7
End Sub
Private Sub TB14_Change()
'On Error Resume Next
Call frm102_calc11
End Sub
Private Sub TB15_Change()
'On Error Resume Next
Call frm102_calc11
End Sub
Private Sub TB16_Change()
'On Error Resume Next
Call frm102_calc11
Call frm102_calc12
Call frm102_calc13
End Sub
Private Sub TB17_Change()
'On Error Resume Next
Call frm102_calc13
End Sub
Private Sub TB19_Change()
'On Error Resume Next
Call frm102_calc11
Call frm102_calc14
Call frm102_calc15
End Sub
Private Sub TB20_Change()
'On Error Resume Next
Call frm102_calc15
End Sub
Private Sub TB22_Change()
'On Error Resume Next
Call frm102_calc11
End Sub
Private Sub TB24_Change()
'On Error Resume Next
Call frm102_calc6
End Sub
Private Sub TB3_Change()
'On Error Resume Next
Call frm102_calc1
End Sub
Private Sub TB4_Change()
'On Error Resume Next
Call frm102_calc2
End Sub
Private Sub TB5_Change()
'On Error Resume Next
Call frm102_calc3
End Sub
Private Sub TB7_Change()
'On Error Resume Next
Call frm102_calc10
End Sub
Private Sub TB8_Change()
'On Error Resume Next
Call frm102_calc10
End Sub
Private Sub TB9_Change()
'On Error Resume Next
If IsNumeric(Frm102.TB9) Then
    Frm102.L13_Text = Format(Frm102.TB9, "#,##0.00") 'Overall : Upah + GST
Else
    Frm102.L13_Text = Format(0, "#,##0.00") 'Overall : Upah + GST
End If

Call frm102_calc10
End Sub
Private Sub Tmr2_Timer()
'On Error Resume Next
Dim Frm102_LM_LIMIT As Integer
Dim Frm102_LM_BIL As Integer

If Frm102.CB1 = 1 And Frm102.TB1 <> vbNullString And Frm102.Tmr2.Enabled = True Then
    If Frm102.Tmr2.Interval = 100 Then
        If InStr(1, Frm102.TB1, "'") <> 0 Then
            MsgBox "No. Siri Produk Mengandungi Simbol Yang Tidak Sah , ['].", vbInformation, "Info"
            
            Frm102.TB1 = vbNullString
            Exit Sub
        End If
        
        Call frm102_reset_1
        Call Frm102_Call_Product_Detail
        
    End If
End If
End Sub
