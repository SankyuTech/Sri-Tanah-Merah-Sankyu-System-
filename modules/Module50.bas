Attribute VB_Name = "Module50"
Sub frm130_kira_jumlah_bayaran()
'on error resume next
Dim Frm130_LM_TUNAI As Double
Dim Frm130_LM_BANK As Double
Dim Frm130_LM_KREDIT As Double
Dim Frm130_LM_SIMPANAN As Double

Frm130_LM_TUNAI = 0
Frm130_LM_BANK = 0
Frm130_LM_KREDIT = 0
Frm130_LM_SIMPANAN = 0

If IsNumeric(frm130.TB27) Then
    Frm130_LM_TUNAI = frm130.TB27
End If
If IsNumeric(frm130.TB28) Then
    Frm130_LM_BANK = frm130.TB28
End If
If IsNumeric(frm130.TB29) Then
    Frm130_LM_KREDIT = frm130.TB29
End If
If IsNumeric(frm130.TB21) Then
    Frm130_LM_SIMPANAN = frm130.TB21
End If

frm130.TB32 = Format(Frm130_LM_TUNAI + Frm130_LM_BANK + Frm130_LM_KREDIT + Frm130_LM_SIMPANAN, "#,##0.00")  'Jumlah Bayaran Keseluruhan
End Sub
Sub Frm130_kira_caj_kad_kredit()
'on error resume next
Dim LM_KAD_KREDIT As Double
Dim LM_CAJ As Double

LM_KAD_KREDIT = 0
LM_CAJ = 0

If ((frm130.TB29 <> vbNullString And IsNumeric(frm130.TB29)) And (frm130.L31_Text <> vbNullString And IsNumeric(frm130.L31_Text))) Then
    LM_KAD_KREDIT = frm130.TB29
    LM_CAJ = frm130.L31_Text
    
    frm130.L32_Text = Format(LM_KAD_KREDIT * (LM_CAJ / 100), "#,##0.00")
Else
    frm130.L32_Text = "0.00"
End If
End Sub
Sub Frm130_kira_caj_gst_kad_kredit()
'on error resume next
Dim LM_CAJ As Double
Dim LM_RATE_GST As Double

LM_CAJ = 0
LM_RATE_GST = 0

If ((frm130.L32_Text <> vbNullString And IsNumeric(frm130.L32_Text)) And (frm130.L8_Text <> vbNullString And IsNumeric(frm130.L8_Text))) Then
    LM_CAJ = frm130.L32_Text
    LM_RATE_GST = frm130.L8_Text
    
    frm130.L81_Text = Format(LM_CAJ * (LM_RATE_GST / 100), "#,##0.00")
Else
    frm130.L81_Text = "0.00"
End If
End Sub
Sub Frm130_kira_potongan_kad_kredit()
'on error resume next
Dim LM_KAD_KREDIT As Double
Dim LM_CAJ As Double
Dim LM_GST As Double

LM_KAD_KREDIT = 0
LM_CAJ = 0
LM_GST = 0

If ((frm130.TB29 <> vbNullString And IsNumeric(frm130.TB29)) And (frm130.L32_Text <> vbNullString And IsNumeric(frm130.L32_Text)) And (frm130.L81_Text <> vbNullString And IsNumeric(frm130.L81_Text))) Then
    LM_KAD_KREDIT = frm130.TB29
    LM_CAJ = frm130.L32_Text
    LM_GST = frm130.L81_Text
    
    frm130.L82_Text = Format(LM_KAD_KREDIT + LM_CAJ + LM_GST, "#,##0.00")
Else
    frm130.L82_Text = "0.00"
End If
End Sub
Sub frm130_initial_setting()
'on error resume next
frm130.TB27 = "0.00"
frm130.TB28 = "0.00"
frm130.TB29 = "0.00"
frm130.TB21 = "0.00"
frm130.TB32 = "0.00"
frm130.TB33 = "0.00"

frm130.L26_Text = "0.00"
frm130.L31_Text = "0.00"
frm130.L32_Text = "0.00"
frm130.L81_Text = "0.00"
frm130.L82_Text = "0.00"

frm130.L41_Text = "0"

frm130.L8_Text = G_RATE_GST

frm130.CBB2.Clear

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 74_cas_kad_kredit where status = 1 order by jenis_kad ASC", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    If Not IsNull(rs!jenis_kad) Then frm130.CBB2.AddItem rs!jenis_kad
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing
End Sub
Sub frm130_reset()
'on error resume next
frm130.TB27 = "0.00"
frm130.TB28 = "0.00"
frm130.TB29 = "0.00"
frm130.TB21 = "0.00"
frm130.TB32 = "0.00"
frm130.TB33 = "0.00"
End Sub
Sub frm130_kiraan_cara_bayaran()
'on error resume next
Dim Frm130_LM_JUMLAH As Double
Dim Frm130_LM_BANK As Double
Dim Frm130_LM_KREDIT As Double
Dim Frm130_LM_SIMPANAN As Double

Frm130_LM_JUMLAH = 0
Frm130_LM_BANK = 0
Frm130_LM_KREDIT = 0
Frm130_LM_SIMPANAN = 0

If IsNumeric(frm130.TB33) Then
    Frm130_LM_JUMLAH = frm130.TB33
End If
If IsNumeric(frm130.TB28) Then
    Frm130_LM_BANK = frm130.TB28
End If
If IsNumeric(frm130.TB29) Then
    Frm130_LM_KREDIT = frm130.TB29
End If
If IsNumeric(frm130.TB21) Then
    Frm130_LM_SIMPANAN = frm130.TB21
End If

frm130.TB27 = Format(Frm130_LM_JUMLAH - (Frm130_LM_BANK + Frm130_LM_KREDIT + Frm130_LM_SIMPANAN), "#,##0.00")  'Jumlah Bayaran Keseluruhan
End Sub

