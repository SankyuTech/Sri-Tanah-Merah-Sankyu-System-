VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frm131 
   Caption         =   "Menu"
   ClientHeight    =   12960
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   21825
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
   Icon            =   "frm131.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   12960
   ScaleWidth      =   21825
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5400
      Top             =   4560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   25
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm131.frx":0ECA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm131.frx":34A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm131.frx":5A7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm131.frx":8058
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm131.frx":A632
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm131.frx":CC0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm131.frx":F1E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm131.frx":117C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm131.frx":15A1A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm131.frx":26274
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm131.frx":2A4CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm131.frx":3AD28
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm131.frx":3D302
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm131.frx":4DB5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm131.frx":5E3B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm131.frx":6EC10
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm131.frx":7F46A
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm131.frx":836C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm131.frx":93F1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm131.frx":A4778
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm131.frx":A6D52
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm131.frx":A932C
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm131.frx":B9B86
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm131.frx":CA3E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm131.frx":DAC3A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView LV1 
      Height          =   2775
      Left            =   3000
      TabIndex        =   0
      Top             =   120
      Width           =   18345
      _ExtentX        =   32359
      _ExtentY        =   4895
      Arrange         =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Scan Item"
         Object.Width           =   2540
         ImageIndex      =   1
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Maklumat Pembeli"
         Object.Width           =   2540
         ImageIndex      =   2
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Bayaran"
         Object.Width           =   2540
         ImageIndex      =   3
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   2580
      Left            =   120
      Picture         =   "frm131.frx":DD214
      Stretch         =   -1  'True
      Top             =   240
      Width           =   2775
   End
End
Attribute VB_Name = "frm131"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'on error resume next

End Sub
Private Sub LV1_Click()
'on error resume next
LM_KEY = frm131.LV1.SelectedItem.Key

If LM_KEY = "Jualan" Then

    If MDI_frm1.L20_Text = "Semua cawangan" Then
    
        Frm96.CMD2.Visible = True
        Frm96.CMD1.Visible = False
    
        Call Frm96_initial
            
        Frm96.Show vbModal
        
    End If
    
    Call MDI_frm1_unload_all_menu
    Call Frm84_background_color
    
    G_CALC_AUTO = 0
    
    Frm84.CB4 = 1
    Call Frm84_form_load
    'Frm84.CB4 = 1
    'Note = "Sila buat pilihan jenis pengiraan upah." & vbCrLf & _
            vbNullString & vbCrLf & _
            "YES : Upah mengikut tetapan per item" & vbCrLf & _
            "NO  : Upah mengikut berat"
    
    'Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    'If Answer = vbNo Then
    '    G_KIRAAN_UPAH = 0
    '    Frm84.L86_Text = "PENGIRAAN UPAH MENGIKUT BERAT"
    'Else
    '    G_KIRAAN_UPAH = 1
    '    Frm84.L86_Text = "PENGIRAAN UPAH MENGIKUT UPAH PER ITEM"
    'End If
    
    Frm84.L62_Text = "Jualan oleh agen dropship : TIDAK"
    MDI_frm1.L5_Text = 4

ElseIf LM_KEY = "Developer" Then

    Call check_internet_dev
    
    Note = "Sila masukkan Developer Pass."

    LM_PASSWORD = InputBox(Note, "Developer", "")
    
    If StrPtr(LM_PASSWORD) = 0 Then
        Exit Sub
    End If
    
    If LM_PASSWORD = G_DEV_PASS Then
        
        Call MDI_frm1_unload_all_menu
        frm150.Show
    
    Else
    
        MsgBox "Pass Key yang dimasukkan TIDAK BETUL.", vbCritical, "Critical"
        Exit Sub
        
    End If

    frm150.LV1.ListItems.Clear
    
    With frm150.LV1
        Set .SmallIcons = frm150.ImageList1
        Set .Icons = frm150.ImageList1

        .ListItems.Add , "License", "License", 2
        .ListItems.Add , "Setting Invoice", "Setting Invoice", 2
        .ListItems.Add , "Invoice", "Invoice", 2
        .ListItems.Add , "Reset Sistem", "Reset Sistem", 2
        
    End With
    
ElseIf LM_KEY = "Tempahan" Then

    Call MDI_frm1_unload_all_menu
    Call Frm93_background_color
    MDI_frm1.L5_Text = 8
    Frm93.Show

ElseIf LM_KEY = "Penerimaan Stok Baru" Then

    If MDI_frm1.L20_Text = "Semua cawangan" Then
    
        Frm96.CMD2.Visible = True
        Frm96.CMD1.Visible = False
    
        Call Frm96_initial
            
        Frm96.Show vbModal
        
    End If
    
    Call MDI_frm1_unload_all_menu
    
    Frm83.CMD24.Enabled = False
    Frm83.CMD25.Enabled = False

    Frm83.L100_Text.Visible = False
    Frm83.L101_Text.Visible = False
    Frm83.L102_Text.Visible = False
    Frm83.TB40.Visible = False
    Frm83.TB41.Visible = False
    Frm83.TB42.Visible = False
            
    Frm83.CMD10.Visible = True
    Frm83.CMD11.Visible = True
    Frm83.CMD20.Visible = False
    Frm83.CMD21.Visible = False
    
    Frm83.CMD1.Visible = False
    Frm83.CMD6.Visible = False
    Frm83.CMD7.Visible = False
    Frm83.CMD12.Visible = False
    Frm83.CMD13.Visible = False
    Frm83.CMD14.Visible = False
    Frm83.CMD2.Visible = False
    Frm83.CMD5.Visible = False
    Frm83.CMD10.Visible = False
    Frm83.CMD11.Visible = False
    
    Frm83.Frame1.Left = 120
    Frm83.Frame1.Top = 120
    
    Frm83.Frame1.Visible = True
    
    'Call Frm96_background_color
    'Call Frm96_initial
    MDI_frm1.L5_Text = 1
    
    Frm83.CB7 = 1 'Barang baru
    Frm83.CB8 = 0 'Used gold
    
    'ID :
        '1 : Stock In -> Penerimaan stok baru (Barang kemas & permata)
        '2 : Stock In -> Penerimaan stok baru (Gold bar)
        
    If MDI_frm1.L5_Text = 1 Then
        Frm83.CB9 = 1 'Barang kemas / permata
        Frm83.CB10 = 0 'Gold bar
    ElseIf MDI_frm1.L5_Text = 2 Then
        Frm83.CB9 = 0 'Barang kemas / permata
        Frm83.CB10 = 1 'Gold bar
        
        Frm83.CB14 = 0
        Frm83.CB15 = 0
        
        Frm83.CB14.Enabled = False
        Frm83.CB15.Enabled = False
    End If
    
    Call Frm83_background_color
    Call Frm83_form_load
    Call frm83_flag_barang_baru
    
    If MDI_frm1.L5_Text = 2 Then
    
        Call Frm83_mode_gold_bar
        
    End If
    
    Frm83.Show

ElseIf LM_KEY = "Trade In" Or LM_KEY = "Belian Emas Terpakai" Then
    
    If MDI_frm1.L20_Text = "Semua cawangan" Then
    
        Frm96.CMD2.Visible = True
        Frm96.CMD1.Visible = False
    
        Call Frm96_initial
            
        Frm96.Show vbModal
        
    End If
    
    Call MDI_frm1_unload_all_menu
    Call Frm83_background_color
    
    Frm83.CMD24.Enabled = True
    Frm83.CMD25.Enabled = True

    Frm83.L100_Text.Visible = True
    Frm83.L101_Text.Visible = True
    Frm83.L102_Text.Visible = True
    Frm83.TB40.Visible = True
    Frm83.TB41.Visible = True
    Frm83.TB42.Visible = True
            
    Frm83.CMD5.Visible = True
    Frm83.CMD2.Visible = True
    Frm83.CMD22.Visible = False
    Frm83.CMD23.Visible = False
    
    Frm83.Frame1.Left = 1680
    Frm83.Frame1.Top = 120
    
    Frm83.Frame9.Left = 1680
    Frm83.Frame9.Top = 120
    
    Frm83.ListView1.Left = 120
    Frm83.ListView1.Top = 120
    
    Frm83.Frame1.Visible = True
    Frm83.ListView1.Visible = True
    
    Frm83.Label40.Visible = True
    Frm83.L10_Text.Visible = True
    
    Frm83.CBB1.Enabled = False
    Frm83.CBB1.BackColor = &H8000000A
    
    Frm83.ListView1.ListItems.Clear
    
    With Frm83.ListView1
        Set .SmallIcons = Frm83.ImageList1
        Set .Icons = Frm83.ImageList1

        .ListItems.Add , "Data Item", "Data Item", 1
        .ListItems.Add , "Senarai Item", "Senarai Item", 2
        
    End With

    Call Frm83_form_load_trade_in
    Call Frm83_form_load
    Call Frm83_form_load_trade_in

    Call frm83_flag_barang_trade_in
    MDI_frm1.L5_Text = 3

ElseIf LM_KEY = "Servis & Belanja" Then

    Call MDI_frm1_unload_all_menu
    Call Frm92_background_color
    Call frm92_setting_report
    
    MDI_frm1.L5_Text = 10
    Frm92.Show

ElseIf LM_KEY = "Maklumat Pelanggan" Then

    Call Frm68_background_color
    Call MDI_frm1_unload_all_menu
    Call Frm68_background_color
    MDI_frm1.L5_Text = 11
    
    Frm68.Show
    Frm68.L36_Text = 0 '0 : Terus dari menu data pelanggan , 1 : Data pembeli , 2 : Data agen dropship

ElseIf LM_KEY = "Pengeluaran & Kemasukkan Tunai" Then

    Call MDI_frm1_unload_all_menu
    Call Frm100_background_color
    
    Frm100.Show
    MDI_frm1.L5_Text = 22

ElseIf LM_KEY = "Pengurusan Buku Cek" Then

    Call MDI_frm1_unload_all_menu
    Call Frm86_background_color
    
    Call Frm86_Initial_Setting
    Frm86.Show
    MDI_frm1.L5_Text = 24

ElseIf LM_KEY = "E-mail Promosi" Then

    Call MDI_frm1_unload_all_menu
    Call Frm97_background_color
    
    Frm97.Show
    MDI_frm1.L5_Text = 25

ElseIf LM_KEY = "Agihan Stok" Then

    Call MDI_frm1_unload_all_menu
    Call Frm108_background_color
    
    Frm108.Show
    MDI_frm1.L5_Text = 28

ElseIf LM_KEY = "Goods Despatch Note (Per Item)" Then

    Note = "Menu ini adalah dikhususkan bagi hantaran barang/tukaran barang dengan pihak supplier atau agen." & vbCrLf & _
            vbNullString & vbCrLf & _
            "Setiap barang yang ditukarkan perlu dipilih SATU PER SATU dari senarai atau SCAN barang tersebut." & vbCrLf & _
            vbNullString & vbCrLf & _
            "Teruskan?"
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbNo Then
        Exit Sub
    End If
    
    If Answer = vbYes Then
    
        If MDI_frm1.L20_Text = "Semua cawangan" Then
        
            Frm96.CMD2.Visible = True
            Frm96.CMD1.Visible = False
        
            Call Frm96_initial
                
            Frm96.Show vbModal
            
        End If
    
        Call MDI_frm1_unload_all_menu
        Call Frm115_background_color
        
        GLOBAL_DISABLE = 0
        Frm115.TB1 = vbNullString
        
        Call Frm115_reset_1
        Call Frm115_reset_2
        Call Frm115_reset_3
        Call Frm115_reset_main
        Call Frm28_initial
        Call Frm115_reset_main2
        
        Frm115.DTPicker1 = DateTime.Date$
        MDI_frm1.L5_Text = 16
        
        Frm115.L69_Text = -1 'Titik Pencarian Data
        Frm115.L70_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
        Frm115.L67_Text = 0 'Paparan Page ke-xxx
        Frm115.L68_Text = 0
        
        Call frm115_initial_setting_stok
        Call Frm115_Senarai_Jualan_Header
        Call frm115_reset_gdn_list
        
        Frm115.CMD8.Visible = True
        Frm115.CMD9.Visible = True
        Frm115.CMD10.Visible = False
        Frm115.CMD11.Visible = False
        
        Frm115.Picture = MDI_frm1.Picture
        Frm115.Show
        
        Frm115.L32_Text = 0 '0 : Data Baru , 1 : Edit Data
        
        Frm115.TB1.SetFocus
        
    End If

ElseIf LM_KEY = "Goods Despatch Note (Bulk)" Then

    Note = "Menu ini adalah dikhususkan bagi hantaran barang/tukaran barang dengan pihak supplier atau agen." & vbCrLf & _
            vbNullString & vbCrLf & _
            "Barangan ini akan dihantar secara BULK." & vbCrLf & _
            vbNullString & vbCrLf & _
            "Teruskan?"
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbNo Then
        Exit Sub
    End If
    
    If Answer = vbYes Then
    
        If MDI_frm1.L20_Text = "Semua cawangan" Then
        
            Frm96.CMD2.Visible = True
            Frm96.CMD1.Visible = False
        
            Call Frm96_initial
                
            Frm96.Show vbModal
            
        End If
    
        Call MDI_frm1_unload_all_menu
        Call frm123_background_color
        
        Call Frm123_one_time_reset
        Call Frm123_reset_1
        Call Frm123_reset_3
        
        frm123.L69_Text = -1 'Titik Pencarian Data
        frm123.L70_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
        frm123.L67_Text = 0 'Paparan Page ke-xxx
        frm123.L68_Text = 0
        
        Call Frm123_Senarai_Belian_Header
        
        frm123.Show
        frm123.L32_Text = 0 '0 : Data Baru , 1 : Edit Data
        
        GLOBAL_DISABLE = 0
        
        MDI_frm1.L5_Text = 34
        
    End If

ElseIf LM_KEY = "Goods Received Note" Then

    If MDI_frm1.L20_Text = "Semua cawangan" Then
    
        Frm96.CMD2.Visible = True
        Frm96.CMD1.Visible = False
    
        Call Frm96_initial
            
        Frm96.Show vbModal
        
    End If
    
    Call MDI_frm1_unload_all_menu
    Call Frm116_background_color
    
    Call frm116_one_time_reset
    Call frm116_reset_1
    Call Frm116_reset_3
    
    frm116.L69_Text = -1 'Titik Pencarian Data
    frm116.L70_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
    frm116.L67_Text = 0 'Paparan Page ke-xxx
    frm116.L68_Text = 0
    
    Call Frm116_Senarai_Belian_Header
    
    frm116.Show
    frm116.L32_Text = 0 '0 : Data Baru , 1 : Edit Data
    
    GLOBAL_DISABLE = 0
    
    MDI_frm1.L5_Text = 17

ElseIf LM_KEY = "Invoice / Voucher" Then

    If MDI_frm1.L20_Text = "Semua cawangan" Then
    
        Frm96.CMD2.Visible = True
        Frm96.CMD1.Visible = False
    
        Call Frm96_initial
            
        Frm96.Show vbModal
        
    End If
    
    Call MDI_frm1_unload_all_menu
    Call Frm118_background_color
    
    Call frm118_initial_setting
    MDI_frm1.L5_Text = 33

ElseIf LM_KEY = "Update Dulang" Then
    
    Call MDI_frm1_unload_all_menu
    Call Frm132_background_color

    Call frm132_setting
    
ElseIf LM_KEY = "Report Keseluruhan" Then

    Call MDI_frm1_unload_all_menu
    Call Frm101_background_color
    
    Call Frm101_initial_setting
    Frm101.Show
    
    Frm101.Pic1.Visible = True
    Frm101.CB2 = 1
    MDI_frm1.L5_Text = 12

ElseIf LM_KEY = "Report Trade In" Then

    Call MDI_frm1_unload_all_menu
    Call Frm129_background_color
    
    Call frm129_initial_setting
    
    frm129.Show

ElseIf LM_KEY = "GDN/GRN" Then

    Call MDI_frm1_unload_all_menu
    Call Frm117_background_color
    Call frm117_pic_ena_disable
    Call frm117_initial_setting
    
    MDI_frm1.L5_Text = 32
    
    frm117.L69_Text = -1 'Titik Pencarian Data
    frm117.L70_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
    frm117.L67_Text = 0 'Paparan Page ke-xxx
    frm117.L68_Text = 0

ElseIf LM_KEY = "Report Kewangan" Then

    Call MDI_frm1_unload_all_menu
    Call Frm105_background_color
    Frm106.Picture = MDI_frm1.Picture
    
    Frm105.Show
    MDI_frm1.L5_Text = 13

ElseIf LM_KEY = "Senarai Invoice" Then

    Call MDI_frm1_unload_all_menu
    Call Frm110_background_color
    
    Frm110.Show
    MDI_frm1.L5_Text = 30

ElseIf LM_KEY = "Penyata Untung Rugi (Restock)" Then

    Call MDI_frm1_unload_all_menu
    Call Frm104_background_color
    
    Frm104.Show
    MDI_frm1.L5_Text = 14

ElseIf LM_KEY = "Penyata Untung Rugi (Runcit)" Then

    Call MDI_frm1_unload_all_menu
    'Call Frm103_background_color
    Call Frm103_initial_setting2
    
    Frm103.CBB1.Clear
    
    Frm103.CBB1.AddItem "Semua cawangan"
    
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from 56_maklumat_kedai where status='" & 1 & "' order by cawangan ASC", cn, adOpenKeyset, adLockOptimistic
    
    While rs.EOF = False
        If Not IsNull(rs!cawangan) Then Frm103.CBB1.AddItem rs!cawangan
        rs.MoveNext
    Wend
    
    rs.Close
    Set rs = Nothing
    
    Frm103.CBB1 = "Semua cawangan"
    
    Frm103.CBB2.Clear
    
    Frm103.CBB2.AddItem "Semua dulang"
    
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from setting_database where status = 1 AND SenaraiDulang is not NULL", cn, adOpenKeyset, adLockOptimistic
    
    While rs.EOF = False
        If Not IsNull(rs!SenaraiDulang) Then Frm103.CBB2.AddItem rs!SenaraiDulang 'Senarai Dulang
        rs.MoveNext
    Wend
    
    rs.Close
    Set rs = Nothing
    
    Frm103.CBB2 = "Semua dulang"
    
    If MDI_frm1.L4_Text <> "HQ" And MDI_frm1.L4_Text <> "Developer" Then
    
        Frm103.CBB1 = MDI_frm1.L20_Text
        Frm103.CBB1.Enabled = False
        
    Else
        
        Frm103.CBB1.Enabled = True
        
    End If
    
    Frm103.Show
    MDI_frm1.L5_Text = 15

ElseIf LM_KEY = "Inventori Dulang" Then

    Call MDI_frm1_unload_all_menu
    Call Frm57_background_color
    
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from setting_database", cn, adOpenKeyset, adLockOptimistic
    
    While rs.EOF = False
        If rs!SenaraiDulang <> vbNullString And rs!Status = 1 Then
            Frm57.CBB1.AddItem rs!SenaraiDulang 'Senarai Dulang
        End If
        rs.MoveNext
    Wend
    
    rs.Close
    Set rs = Nothing
    
    Call Frm57_M_Clear
    
    Frm57.Show
    Frm57.Pic1.Visible = True
    
    MDI_frm1.L5_Text = 21

ElseIf LM_KEY = "Report GST" Then

    Call MDI_frm1_unload_all_menu
    Call Frm75_background_color
    
    Frm75.L17_Text = "0.00"
    Frm75.L18_Text = "0.00"
    Frm75.L19_Text = vbNullString
    
    Frm75.L69_Text = -1 'Titik Pencarian Data
    Frm75.L75_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
    Frm75.L67_Text = 0 'Paparan Page ke-xxx
    Frm75.L68_Text = 0
    
    Frm75.L62_Text = -1 'Start Point
    Frm75.L60_Text = 0 'Current Page
    Frm75.L61_Text = 0 'Current Page
    Frm75.L63_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
        
    Frm75.Show
    MDI_frm1.L5_Text = 15

ElseIf LM_KEY = "Barang Hilang" Then

    Call MDI_frm1_unload_all_menu
    Call Frm126_background_color
    
    frm126.L69_Text = -1 'Titik Pencarian Data
    frm126.L70_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
    frm126.L67_Text = 0 'Paparan Page ke-xxx
    frm126.L68_Text = 0
    
    frm126.L10_Text = 0
    frm126.L11_Text = "0.00 g"
    frm126.L12_Text = "RM 0.00"
    
    frm126.Show

ElseIf LM_KEY = "Log" Then

    Call MDI_frm1_unload_all_menu
    Call Frm127_background_color
    
    frm127.CB1 = 0
    frm127.DTPicker1 = DateTime.Date
    frm127.DTPicker2 = DateTime.Date
    
    frm127.TB1 = vbNullString
    
    frm127.CBB1.Clear
    frm127.CBB2.Clear
    frm127.CBB3.Clear
    
    frm127.CBB1.AddItem "Semua terminal"
    
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from 91_senarai_terminal where status = 1 order by terminal ASC", cn, adOpenKeyset, adLockOptimistic
    
    While rs.EOF = False

        If Not IsNull(rs!terminal) Then frm127.CBB1.AddItem rs!terminal

        rs.MoveNext
    Wend
    
    rs.Close
    Set rs = Nothing
    
    frm127.CBB1 = "Semua terminal"
    
    frm127.CBB3.AddItem "Semua user"
    
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from employee where status='" & "Aktif" & "' AND InvestorSmall = 0 AND InvestorBig = 0 order by samaran ASC", cn, adOpenKeyset, adLockOptimistic
    
    While rs.EOF = False
        
        If rs!user_level = 1 Or rs!user_level = 2 Or rs!user_level = 3 Or rs!user_level = G_LEVEL_USER Then frm127.CBB3.AddItem rs!Samaran
        rs.MoveNext
    Wend
    
    rs.Close
    Set rs = Nothing
    
    frm127.CBB3 = "Semua user"
    
    frm127.CBB2.AddItem "Semua cawangan"
    
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from 56_maklumat_kedai where status='" & 1 & "' order by cawangan ASC", cn, adOpenKeyset, adLockOptimistic
    
    While rs.EOF = False
        If Not IsNull(rs!cawangan) Then frm127.CBB2.AddItem rs!cawangan
        rs.MoveNext
    Wend
    
    rs.Close
    Set rs = Nothing
    
    frm127.CBB2 = "Semua cawangan"
    
    If MDI_frm1.L4_Text <> "HQ" And MDI_frm1.L4_Text <> "Developer" Then
    
        frm127.CBB2 = MDI_frm1.L20_Text
        frm127.CBB2.Enabled = False
        
    Else
        
        frm127.CBB2.Enabled = True
        
    End If
    
    frm127.L69_Text = -1 'Titik Pencarian Data
    frm127.L70_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
    frm127.L67_Text = 0 'Paparan Page ke-xxx
    frm127.L68_Text = 0
    
    frm127.Show
    
ElseIf LM_KEY = "Report Stok Dulang" Then
    
    If MDI_frm1.L20_Text = "Semua cawangan" Then
    
        Frm96.CMD2.Visible = True
        Frm96.CMD1.Visible = False
    
        Call Frm96_initial
            
        Frm96.Show vbModal
        
    End If

    Note = "Sistem mungkin akan mengambil masa untuk mengeluarkan report ini." & vbCrLf & _
            vbNullString & vbCrLf & _
            "*** SILA PASTIKAN ANDA HANYA MENGELUARKAN REPORT INI DARI SATU STATION (KOMPUTER) SAHAJA ***" & vbCrLf & _
            vbNullString & vbCrLf & _
            "Teruskan?"
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbYes Then
            
        frm134.L69_Text = -1 'Titik Pencarian Data
        frm134.L70_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
        frm134.L67_Text = 0 'Paparan Page ke-xxx
        frm134.L68_Text = 0
        
        Call frm134_report_stok_dulang
        
        GM_NEXT_PREV = 0
        
        Call frm134_report_stok_header
        Call frm134_report_stok
        
        frm134.Show
        
    End If
    
ElseIf LM_KEY = "Tetapan Asas Sistem" Then

    Call MDI_frm1_unload_all_menu
    Call Frm95_background_color
    
    MDI_frm1.L5_Text = 18

ElseIf LM_KEY = "Tetapan Harga Jualan Emas" Then

    Call MDI_frm1_unload_all_menu
    Call Frm54_background_color
    
    Frm54.CBB1.Clear
    
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from setting_database order by Kod_Metal_Purity DESC", cn, adOpenKeyset, adLockOptimistic
    
    While rs.EOF = False
        If rs!Metal_Purity <> vbNullString Then
            Frm54.CBB1.AddItem rs!Kod_Metal_Purity
        End If
        rs.MoveNext
    Wend
    
    rs.Close
    Set rs = Nothing
    
    Frm54.Show
    
    MDI_frm1.L5_Text = 19

ElseIf LM_KEY = "Tetapan Sistem" Then

    Call MDI_frm1_unload_all_menu
    Call Frm111_background_color
    
    Call Frm111_initial_setting
    Call Frm111_setting
    
    Call sys_config_membership
    
    If G_MODE = "NO" Then
        Frm111.L4_Text.Visible = False
    Else
        Frm111.L4_Text.Visible = True
    End If
    
    Frm111.Show
    MDI_frm1.L5_Text = 30

ElseIf LM_KEY = "Data Pekerja" Then
    
    Call MDI_frm1_unload_all_menu
    
    Frm49.LV1.ListItems.Clear
    
    With Frm49.LV1
        Set .SmallIcons = Frm49.ImageList1
        Set .Icons = Frm49.ImageList1

        .ListItems.Add , "Pendaftaran Pekerja", "Pendaftaran Pekerja", 1
        .ListItems.Add , "Senarai Pekerja", "Senarai Pekerja", 2
        
    End With
    
    Call frm49_disable_form
    Call Frm49_background_color
    Call frm49_Default
    Call frm49_cawangan
    
    Frm49.Show
    Frm49.Frame1.Visible = True
    MDI_frm1.L5_Text = 17
    
    Frm49.TB1.SetFocus

ElseIf LM_KEY = "Tetapan Barcode" Then

    If MDI_frm1.L20_Text = "Semua cawangan" Then
    
        Frm96.CMD2.Visible = True
        Frm96.CMD1.Visible = False
    
        Call Frm96_initial
            
        Frm96.Show vbModal
        
    End If
    
    Call MDI_frm1_unload_all_menu
    Frm56.Picture = MDI_frm1.Picture
    
    Frm56.Show
    MDI_frm1.L5_Text = 20

ElseIf LM_KEY = "Payroll" Then

    Call MDI_frm1_unload_all_menu
    Call Frm48_background_color
    
    Call Frm48_Default
    Frm48.Show
    MDI_frm1.L5_Text = 23

ElseIf LM_KEY = "Analisa Harga Emas" Then

    Call MDI_frm1_unload_all_menu
    Frm109.Show
    MDI_frm1.L5_Text = 29

ElseIf LM_KEY = "Backup Database" Then

    Note = "Adakah anda ingin backup database ini?" & vbCrLf & _
            "Sistem tidak dapat beroperasi semasa database dibackup sehingga selesai." & vbCrLf & _
            vbNullString & vbCrLf & _
            "Teruskan?"
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbYes Then
            
        Call backup_database
        
    End If

End If

MDI_frm1.CMD44.Enabled = False
            
End Sub
