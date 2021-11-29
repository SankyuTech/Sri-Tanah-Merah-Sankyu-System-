Attribute VB_Name = "mod_recovery"
Sub recovery_senarai_pelanggan()
'on error resume next
Dim rs10 As ADODB.Recordset

Set rs10 = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

strsql = "insert into " & G_RECOVERY_DATABASE & ".senarai_pelanggan" & "(id_asal,kategori_pelanggan,nama,no_ic,no_tel,email,alamat,nama_waris,no_tel_waris,alamat_waris,nama_bank,nama_akaun,no_akaun,write_timestamp,no_pelanggan,baki_simpanan,dropship,membership_card,yuran_flag,jumlah_yuran,tarikh,no_invoice,status,baki_point,kategori_asal,no_staff,terminal,jenis_urusan,remarks)" & _
            "select ID,kategori_pelanggan,nama,no_ic,no_tel,email,alamat,nama_waris,no_tel_waris,alamat_waris,nama_bank,nama_akaun,no_akaun,write_timestamp,no_pelanggan,baki_simpanan,dropship,membership_card,yuran_flag,jumlah_yuran,tarikh,no_invoice,status,baki_point,kategori_asal,no_staff,terminal,jenis_urusan,remarks from " & G_SERVER_DATABASE & ".senarai_pelanggan WHERE ID='" & G_ID & "'"

Set rs10 = cn.Execute(strsql)
Set rs10 = Nothing
End Sub
Sub recovery_16_gold_bar_belian()
'on error resume next
Dim rs10 As ADODB.Recordset

Set rs10 = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

strsql = "insert into " & G_RECOVERY_DATABASE & ".16_gold_bar_belian" & "(id_asal,no_rujukan,tarikh,cara_bayaran,tunai,bank_in,kad_kredit,cas_kad_kredit,jumlah_kad_kredit,kad_debit,cas_kad_debit,jumlah_kad_debit,cheque,no_cheque,jumlah_asal,diskaun_percent,jumlah_diskaun,gst_zr_harga,gst_zr_cukai,gst_sr_harga,gst_sr_cukai,no_id_gst_supplier,no_resit_supplier,kod_supplier,jumlah_tanpa_gst,jumlah_dengan_gst,gst_ari_nashi,gst_value,jumlah_gst,jumlah_sebenar,flag_trade_in,trade_in_status,no_resit_trade_in,no_pekerja,no_rujukan_pelanggan_buyback,kategori_penjual,status,terminal,write_timestamp,write_timestamp2,remarks,no_staff,jenis_urusan)" & _
            "select ID,no_rujukan,tarikh,cara_bayaran,tunai,bank_in,kad_kredit,cas_kad_kredit,jumlah_kad_kredit,kad_debit,cas_kad_debit,jumlah_kad_debit,cheque,no_cheque,jumlah_asal,diskaun_percent,jumlah_diskaun,gst_zr_harga,gst_zr_cukai,gst_sr_harga,gst_sr_cukai,no_id_gst_supplier,no_resit_supplier,kod_supplier,jumlah_tanpa_gst,jumlah_dengan_gst,gst_ari_nashi,gst_value,jumlah_gst,jumlah_sebenar,flag_trade_in,trade_in_status,no_resit_trade_in,no_pekerja,no_rujukan_pelanggan_buyback,kategori_penjual,status,terminal,write_timestamp,write_timestamp2,remarks,no_staff,jenis_urusan from " & G_SERVER_DATABASE & ".16_gold_bar_belian WHERE ID='" & G_ID & "'"

Set rs10 = cn.Execute(strsql)
Set rs10 = Nothing
End Sub
Sub recovery_data_database()
'on error resume next
Dim rs10 As ADODB.Recordset

Set rs10 = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

strsql = "insert into " & G_RECOVERY_DATABASE & ".data_database" & "(id_asal,NoRujukanSistem,supplier_ID,nama_Supplier,Kod_Supplier,purity_ID,Purity,kod_Purity,kategori_produk_ID,kategori_Produk,kod_Kategori_Produk," _
            & "no_siri_Produk,Berat,Beza_Berat,Upah,Upah30,Upah_Jualan,kos_Belian_Gram,kos_Belian_Item,cara_Belian,tarikh_Belian," _
            & "dimension_Panjang,dimension_Lebar,dimension_Dia,harga_Per_Gram_Item,receiving_Status,code_Supplier,tarikh_Jualan1,harga_lepas_spread," _
            & "SpreadValue,adjustment,bill_No_Trade_In,bill_No_Belian,Cara_Jualan,Nama,No_IC,No_Passport,No_HP,Email,Barcode,Dulang,Market,Resit,StatusItem," _
            & "Upah_Member,KategoriUpah,Upah_Pengedar,Upah_Stokis,upah_normal_dealer,upah_master_dealer,HargaJualan_Member,HargaJualan_Pengedar," _
            & "HargaJualan_RAF,hargajualan_normal_dealer,hargajualan_master_dealer,Upah_RAF,write_timestamp,write_timestamp2,riyal,kos_item_tanpa_tax,gst_ari_nashi," _
            & "kadar_gst,jumlah_gst,harga_item,remarks,no_rujukan_pelanggan_buyback,flag_image,kod_upah,no_cert,harga_tanpa_gst,gst_included,gst_barang_atau_upah," _
            & "jenis_trade_in,form_out_status,cawangan_id,no_rujukan_pulang,flag_upah,upah_per_gram,no_id_gst,terminal,susut_berat,no_pekerja,menu)" & _
            "select ID,NoRujukanSistem,supplier_ID,nama_Supplier,Kod_Supplier,purity_ID,Purity,kod_Purity,kategori_produk_ID,kategori_Produk,kod_Kategori_Produk," _
            & "no_siri_Produk,Berat,Beza_Berat,Upah,Upah30,Upah_Jualan,kos_Belian_Gram,kos_Belian_Item,cara_Belian,tarikh_Belian," _
            & "dimension_Panjang,dimension_Lebar,dimension_Dia,harga_Per_Gram_Item,receiving_Status,code_Supplier,tarikh_Jualan1,harga_lepas_spread," _
            & "SpreadValue,adjustment,bill_No_Trade_In,bill_No_Belian,Cara_Jualan,Nama,No_IC,No_Passport,No_HP,Email,Barcode,Dulang,Market,Resit,StatusItem," _
            & "Upah_Member,KategoriUpah,Upah_Pengedar,Upah_Stokis,upah_normal_dealer,upah_master_dealer,HargaJualan_Member,HargaJualan_Pengedar," _
            & "HargaJualan_RAF,hargajualan_normal_dealer,hargajualan_master_dealer,Upah_RAF,write_timestamp,write_timestamp2,riyal,kos_item_tanpa_tax,gst_ari_nashi," _
            & "kadar_gst,jumlah_gst,harga_item,remarks,no_rujukan_pelanggan_buyback,flag_image,kod_upah,no_cert,harga_tanpa_gst,gst_included,gst_barang_atau_upah," _
            & "jenis_trade_in,form_out_status,cawangan_id,no_rujukan_pulang,flag_upah,upah_per_gram,no_id_gst,terminal,susut_berat,no_pekerja,menu " _
            & "from " & G_SERVER_DATABASE & ".data_database WHERE ID='" & G_ID & "'"

Set rs10 = cn.Execute(strsql)
Set rs10 = Nothing
End Sub
Sub recovery_44_senarai_pelanggan()
'on error resume next
Dim rs10 As ADODB.Recordset

Set rs10 = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

strsql = "insert into " & G_RECOVERY_DATABASE & ".44_senarai_pelanggan" & "(id_asal,tarikh,no_resit,nama,no_tel,write_timestamp,no_staff,terminal,jenis_urusan)" & _
            "select ID,tarikh,no_resit,nama,no_tel,write_timestamp,no_staff,terminal,jenis_urusan from " & G_SERVER_DATABASE & ".44_senarai_pelanggan WHERE ID='" & G_ID & "'"

Set rs10 = cn.Execute(strsql)
Set rs10 = Nothing
End Sub
Sub recovery_24_rekod_kewangan_pelanggan()
'on error resume next
Dim rs10 As ADODB.Recordset

Set rs10 = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

strsql = "insert into " & G_RECOVERY_DATABASE & ".24_rekod_kewangan_pelanggan" & "(id_asal,tarikh,jenis,no_rujukan_pelanggan,no_resit,jumlah,jenis_penggunaan,no_rujukan_pekerja,write_timestamp,terminal,jenis_urusan)" & _
            "select ID,tarikh,jenis,no_rujukan_pelanggan,no_resit,jumlah,jenis_penggunaan,no_rujukan_pekerja,write_timestamp,terminal,jenis_urusan from " & G_SERVER_DATABASE & ".24_rekod_kewangan_pelanggan WHERE ID='" & G_ID & "'"

Set rs10 = cn.Execute(strsql)
Set rs10 = Nothing
End Sub
Sub recovery_22_jualan()
'on error resume next
Dim rs10 As ADODB.Recordset

Set rs10 = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

strsql = "insert into " & G_RECOVERY_DATABASE & ".22_jualan" & "(id_asal,no_resit,tarikh,tunai,bank_in,kad_kredit,jenis_kad,cas_kad_kredit,jumlah_cas_kad_kredit,kadar_gst_kad_kredit," _
            & "gst_kad_kredit,jumlah_potongan_kad_kredit,duit_simpanan_kedai,kad_debit,cas_kad_debit,jumlah_cas_kad_debit," _
            & "jumlah_potongan_kad_debit,jumlah_bayaran,harga_barang,jumlah_cukai_gst,harga_barang_dengan_gst,diskaun," _
            & "harga_lepas_diskaun,adjustment,harga_jualan,loss_trade_in,loss_trade_in_rm,flag_bayaran,jumlah_perlu_bayar," _
            & "kuantiti_barang,jumlah_berat,gst_zr_harga,gst_zr_cukai,gst_sr_harga,gst_sr_cukai,no_pekerja,no_rujukan_pembeli," _
            & "no_rujukan_agen_dropship,flag_trade_in,no_resit_trade_in,jumlah_trade_in,kategori_pembeli,jualan_online," _
            & "write_timestamp,write_timestamp2,bonus_kadar,bonus_jumlah,bonus_pemalar_916,bonus_berat,redeem_flag,redeem_baki_bayaran," _
            & "redeem_berat_bonus_terkumpul,redeem_berat,redeem_harga_semasa916,redeem_jumlah,redeem_baki_bayaran_akhir,potongan_koperasi," _
            & "no_aggrement,no_borang,no_approval,lokasi,invoice_type,epp,approval_code_epp,caj_pos,jenis_trade_in,status,no_tracking," _
            & "bil_rasmi,point_ari_nashi,jumlah_point,kupon_diskaun,kadar_diskaun,kadar_peroleh_point,kadar_tebus_point,terminal,Menu,no_staff)" & _
            "select ID,no_resit,tarikh,tunai,bank_in,kad_kredit,jenis_kad,cas_kad_kredit,jumlah_cas_kad_kredit,kadar_gst_kad_kredit," _
            & "gst_kad_kredit,jumlah_potongan_kad_kredit,duit_simpanan_kedai,kad_debit,cas_kad_debit,jumlah_cas_kad_debit," _
            & "jumlah_potongan_kad_debit,jumlah_bayaran,harga_barang,jumlah_cukai_gst,harga_barang_dengan_gst,diskaun," _
            & "harga_lepas_diskaun,adjustment,harga_jualan,loss_trade_in,loss_trade_in_rm,flag_bayaran,jumlah_perlu_bayar," _
            & "kuantiti_barang,jumlah_berat,gst_zr_harga,gst_zr_cukai,gst_sr_harga,gst_sr_cukai,no_pekerja,no_rujukan_pembeli," _
            & "no_rujukan_agen_dropship,flag_trade_in,no_resit_trade_in,jumlah_trade_in,kategori_pembeli,jualan_online," _
            & "write_timestamp,write_timestamp2,bonus_kadar,bonus_jumlah,bonus_pemalar_916,bonus_berat,redeem_flag,redeem_baki_bayaran," _
            & "redeem_berat_bonus_terkumpul,redeem_berat,redeem_harga_semasa916,redeem_jumlah,redeem_baki_bayaran_akhir,potongan_koperasi," _
            & "no_aggrement,no_borang,no_approval,lokasi,invoice_type,epp,approval_code_epp,caj_pos,jenis_trade_in,status,no_tracking," _
            & "bil_rasmi,point_ari_nashi,jumlah_point,kupon_diskaun,kadar_diskaun,kadar_peroleh_point,kadar_tebus_point,terminal,Menu,no_staff " _
            & "from " & G_SERVER_DATABASE & ".22_jualan WHERE ID='" & G_ID & "'"

Set rs10 = cn.Execute(strsql)
Set rs10 = Nothing
End Sub
Sub recovery_71_tebus_agih_point()
'on error resume next
Dim rs10 As ADODB.Recordset

Set rs10 = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

strsql = "insert into " & G_RECOVERY_DATABASE & ".71_tebus_agih_point" & "(id_asal,no_invoice,tarikh,no_ahli,harga_layak_bonus,kadar_peroleh_point,jumlah_tebus_point," _
            & "kadar_tebus_point,nilaian_tebus_point,write_timestamp,write_timestamp2,write_timestamp3,status," _
            & "remarks,type,no_pekerja,bil_rasmi,terminal,jenis_urusan)" & _
            "select ID,no_invoice,tarikh,no_ahli,harga_layak_bonus,kadar_peroleh_point,jumlah_tebus_point," _
            & "kadar_tebus_point,nilaian_tebus_point,write_timestamp,write_timestamp2,write_timestamp3,status," _
            & "remarks,type,no_pekerja,bil_rasmi,terminal,jenis_urusan " _
            & "from " & G_SERVER_DATABASE & ".71_tebus_agih_point WHERE ID='" & G_ID & "'"

Set rs10 = cn.Execute(strsql)
Set rs10 = Nothing
End Sub
Sub recovery_23_senarai_jualan()
'on error resume next
Dim rs10 As ADODB.Recordset

Set rs10 = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

strsql = "insert into " & G_RECOVERY_DATABASE & ".23_senarai_jualan" & "(id_asal,flag_barang,nama_purity,no_invoice_r,bil_rasmi,modal_tanpa_gst,harga_per_gram_tanpa_gst,jualan_per_gram_dengan_gst,status_r,tarikh,no_resit,no_siri_produk,kategori_produk,purity,berat_asal,berat_jualan,harga_semasa,upah," _
            & "harga_asal,diskaun,harga_lepas_diskaun,adjustment,harga_jualan,gst_ari_nashi,kadar_gst,jumlah_gst," _
            & "harga_dengan_gst,no_pekerja,no_rujukan_pembeli,dropship,no_rujukan_agen_dropship,komisyen_per_gram," _
            & "jumlah_komisyen,status,type,potong_flag,harga_per_gram_modal,modal,untung,harga_per_gram_supplier," _
            & "untung2,kategori_pembeli,dulang,jualan_online,write_timestamp,write_timestamp2,gst_include," _
            & "harga_tanpa_gst,harga_koperasi,lokasi,kadar_penurunan_upah,kadar_penurunan_bp,harga_semasa_staff," _
            & "harga_bp_asal,upah_asal,harga_staff,komisyen_staff,pemalar_tukaran_999,berat_999,jenis_jualan," _
            & "upah_modal,gst_barang_atau_upah,harga_jualan_dengan_gst,kadar_komisyen_upah,komisyen_upah," _
            & "status_rekod,jualan_per_gram,modal_per_gram,flag_upah,upah_per_gram,jenis_urusan,terminal,no_staff)" & _
            "select ID,flag_barang,nama_purity,no_invoice_r,bil_rasmi,modal_tanpa_gst,harga_per_gram_tanpa_gst,jualan_per_gram_dengan_gst,status_r,tarikh,no_resit,no_siri_produk,kategori_produk,purity,berat_asal,berat_jualan,harga_semasa,upah," _
            & "harga_asal,diskaun,harga_lepas_diskaun,adjustment,harga_jualan,gst_ari_nashi,kadar_gst,jumlah_gst," _
            & "harga_dengan_gst,no_pekerja,no_rujukan_pembeli,dropship,no_rujukan_agen_dropship,komisyen_per_gram," _
            & "jumlah_komisyen,status,type,potong_flag,harga_per_gram_modal,modal,untung,harga_per_gram_supplier," _
            & "untung2,kategori_pembeli,dulang,jualan_online,write_timestamp,write_timestamp2,gst_include," _
            & "harga_tanpa_gst,harga_koperasi,lokasi,kadar_penurunan_upah,kadar_penurunan_bp,harga_semasa_staff," _
            & "harga_bp_asal,upah_asal,harga_staff,komisyen_staff,pemalar_tukaran_999,berat_999,jenis_jualan," _
            & "upah_modal,gst_barang_atau_upah,harga_jualan_dengan_gst,kadar_komisyen_upah,komisyen_upah," _
            & "status_rekod,jualan_per_gram,modal_per_gram,flag_upah,upah_per_gram,jenis_urusan,terminal,no_staff " _
            & "from " & G_SERVER_DATABASE & ".23_senarai_jualan WHERE ID='" & G_ID & "'"

Set rs10 = cn.Execute(strsql)
Set rs10 = Nothing
End Sub
Sub recovery_40_tempahan_deposit()
'on error resume next
Dim rs10 As ADODB.Recordset

Set rs10 = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

strsql = "insert into " & G_RECOVERY_DATABASE & ".40_tempahan_deposit" & "(id_asal,no_rujukan_tempahan,no_resit_tempahan,jenis_tempahan,type_barang_kemas,no_siri_produk,kategori_produk," _
            & "purity,anggaran_berat,berat_asal,berat_jualan,harga_semasa,upah,harga_asal,adjustment,anggaran_harga," _
            & "status,kategori_pembeli,nama,no_ic,no_tel,no_rujukan_pelanggan,flag_trade_in,no_resit_trade_in," _
            & "nilaian_trade_in,jumlah_deposit_tunai,jumlah_deposit_trade_in,jumlah_tanpa_gst,jumlah_dengan_gst," _
            & "jumlah_perlu_bayar,adjustment_bayaran,jumlah_bayaran,tarikh,remarks,no_pekerja,write_timestamp," _
            & "harga_asal_barang_permata,bil_batu,harga_batu,jumlah_harga_batu,panjang,lebar,saiz,status_tukang,write_timestamp2," _
            & "terminal,status_invoice,bil_rasmi,no_staff)" & _
            "select ID,no_rujukan_tempahan,no_resit_tempahan,jenis_tempahan,type_barang_kemas,no_siri_produk,kategori_produk," _
            & "purity,anggaran_berat,berat_asal,berat_jualan,harga_semasa,upah,harga_asal,adjustment,anggaran_harga," _
            & "status,kategori_pembeli,nama,no_ic,no_tel,no_rujukan_pelanggan,flag_trade_in,no_resit_trade_in," _
            & "nilaian_trade_in,jumlah_deposit_tunai,jumlah_deposit_trade_in,jumlah_tanpa_gst,jumlah_dengan_gst," _
            & "jumlah_perlu_bayar,adjustment_bayaran,jumlah_bayaran,tarikh,remarks,no_pekerja,write_timestamp," _
            & "harga_asal_barang_permata,bil_batu,harga_batu,jumlah_harga_batu,panjang,lebar,saiz,status_tukang,write_timestamp2," _
            & "terminal,status_invoice,bil_rasmi,no_staff " _
            & "from " & G_SERVER_DATABASE & ".40_tempahan_deposit WHERE ID='" & G_ID & "'"

Set rs10 = cn.Execute(strsql)
Set rs10 = Nothing
End Sub
Sub recovery_42_tempahan_siap()
'on error resume next
Dim rs10 As ADODB.Recordset

Set rs10 = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

strsql = "insert into " & G_RECOVERY_DATABASE & ".42_tempahan_siap" & "(id_asal,no_rujukan_tempahan,no_resit_tempahan,flag_bayaran,jenis_tempahan,type_barang_kemas,no_siri_produk," _
            & "kategori_produk,purity,dulang,berat_asal,berat_jualan,harga_semasa,upah,harga_asal,adjustment,harga," _
            & "kategori_pembeli,nama,no_ic,no_tel,no_rujukan_pelanggan,flag_trade_in,no_resit_trade_in,nilaian_trade_in," _
            & "jumlah_harga_jualan,bayaran_sudah_jelas,baki,jumlah_gst,baki_dengan_gst,baki_adjustment,jumlah_baki_terakhir," _
            & "tarikh,no_pekerja,write_timestamp,gst_include,harga_tanpa_gst,terminal,status_invoice," _
            & "bil_rasmi,no_staff,write_timestamp2)" & _
            "select ID,no_rujukan_tempahan,no_resit_tempahan,flag_bayaran,jenis_tempahan,type_barang_kemas,no_siri_produk," _
            & "kategori_produk,purity,dulang,berat_asal,berat_jualan,harga_semasa,upah,harga_asal,adjustment,harga," _
            & "kategori_pembeli,nama,no_ic,no_tel,no_rujukan_pelanggan,flag_trade_in,no_resit_trade_in,nilaian_trade_in," _
            & "jumlah_harga_jualan,bayaran_sudah_jelas,baki,jumlah_gst,baki_dengan_gst,baki_adjustment,jumlah_baki_terakhir," _
            & "tarikh,no_pekerja,write_timestamp,gst_include,harga_tanpa_gst,terminal,status_invoice," _
            & "bil_rasmi,no_staff,write_timestamp2 " _
            & "from " & G_SERVER_DATABASE & ".42_tempahan_siap WHERE ID='" & G_ID & "'"

Set rs10 = cn.Execute(strsql)
Set rs10 = Nothing
End Sub
Sub recovery_77_gdn_grn()
'on error resume next
Dim rs10 As ADODB.Recordset

Set rs10 = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

strsql = "insert into " & G_RECOVERY_DATABASE & ".77_gdn_grn" & "(id_asal,tarikh,masa,write_timestamp,no_rujukan,berat_asal,kadar_tukaran,berat_tukaran," _
            & "no_rujukan_supplier,harga_999,revision,no_rujukan_rev,harga_tanpa_gst,jumlah_gst,kadar_gst,harga_dengan_gst,Status,jenis_urusan,terminal,user,supplier_agen," _
            & "jenis,harga_dengan_gst_grn,harga_tanpa_gst_grn,berat_tukaran_grn,bil_barang,gst_sr_harga,gst_zr_harga,gst_sr_cukai,gst_zr_cukai,nilaian_harga_emas)" & _
            "select ID,tarikh,masa,write_timestamp,no_rujukan,berat_asal,kadar_tukaran,berat_tukaran," _
            & "no_rujukan_supplier,harga_999,revision,no_rujukan_rev,harga_tanpa_gst,jumlah_gst,kadar_gst,harga_dengan_gst,Status,jenis_urusan,terminal,user,supplier_agen," _
            & "jenis,harga_dengan_gst_grn,harga_tanpa_gst_grn,berat_tukaran_grn,bil_barang,gst_sr_harga,gst_zr_harga,gst_sr_cukai,gst_zr_cukai,nilaian_harga_emas " _
            & "from " & G_SERVER_DATABASE & ".77_gdn_grn WHERE ID='" & G_ID & "'"

Set rs10 = cn.Execute(strsql)
Set rs10 = Nothing
End Sub
