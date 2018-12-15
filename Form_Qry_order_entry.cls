VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Qry_order_entry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Antall_AfterUpdate()
Dim cRecord As Integer

On Error GoTo ErrHandler:

cRecord = Form_Frm_alle_ansatte_med_kundenummer.CurrentRecord

Me.Tbl_Kvotefiler_fra_Zalaris_Øl = Me.Total_øl


If Me.Gratis_brus.Value = True Then
    Me.Gratis = Me.Total_brus
    Me.Tbl_Kvotefiler_fra_Zalaris_Brus = 0
Else
    Me.Tbl_Kvotefiler_fra_Zalaris_Brus = Me.Total_brus
    Me.Gratis = 0
End If

Me.Salgs_dato = Now()

If Me.Total_brus > 0 Then
    Me.Gratis_brus.Enabled = True
End If

Form_Frm_alle_ansatte_med_kundenummer.Refresh

If (Form_Frm_alle_kvoter_per_ansatt_sum_siste_3_mnd.Sum_Of_Brus < 0 Or Form_Frm_alle_kvoter_per_ansatt_sum_siste_3_mnd.Sum_Of_Øl < 0 Or Form_Frm_alle_kvoter_per_ansatt_sum_siste_3_mnd.Sum_Of_Gratis < 0) Then
    MsgBox ("Det er for lite kvote")
End If

ErrHandler:
    If Err.Number = 3101 Then
        MsgBox ("Materialnummeret finnes ikke!")
        Me.Materialnummer = 0
        Me.Materialnummer.SetFocus
        Resume
    End If
End Sub


Private Sub Form_Load()
Me.Filter = "(Year([Qry_order_entry].[Salgs_dato])=Year(Date()) AND Month([Qry_order_entry].[Salgs_dato])=Month(Date()))"

Me.FilterOn = True
End Sub

Private Sub Gratis_brus_Click()
Antall_AfterUpdate
End Sub
