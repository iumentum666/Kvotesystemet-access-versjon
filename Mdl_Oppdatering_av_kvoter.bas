Attribute VB_Name = "Mdl_Oppdatering_av_kvoter"
Option Compare Database

Sub importere_fil()

MsgBox ("Velkommen til filoppdatereren. Velg en kvotefil i neste dialogboks.")

Dim fd As FileDialog
Set fd = Application.FileDialog(msoFileDialogFilePicker)

DoCmd.SetWarnings False
DoCmd.RunSQL "DELETE * FROM Tbl_innlasting_av_kvotefiler"


'get the number of the button chosen
Dim FileChosen As Integer
FileChosen = fd.Show

If FileChosen <> -1 Then
    'Melding om at man trykket cancel
    MsgBox "Du valgte cancel"
Else
    'Her gjør vi alle endringene i databasen
    DoCmd.TransferText _
        TransferType:=acImportDelim, _
        SpecificationName:="Kvotefil", _
        TableName:="Tbl_innlasting_av_kvotefiler", _
        FileName:=fd.SelectedItems(1), _
        HasFieldNames:=True
    MsgBox ("Fil: " & vbNewLine & fd.SelectedItems(1) & vbNewLine & " har blitt importert. Systemet vil nå korrigere kvotene.")
    'DoCmd.OpenQuery "Qry_Oppdatere_tabell_med_korrigerte_kvoter", acViewNormal, acEdit
    DoCmd.OpenQuery "Qry_kvotefiler_finne_alle_inactive_og_withdrawn", acViewNormal, acEdit
    DoCmd.OpenQuery "Qry_kvotefiler_legge_inn_c_på_de_som_mangler_kvotekode", acViewNormal, acEdit
    DoCmd.OpenQuery "Qry_kvotefiler_korrigere_ansattdato", acViewNormal, acEdit
    DoCmd.OpenQuery "Qry_kvotefiler_oppdatere_kvotekoder", acViewNormal, acEdit
    DoCmd.OpenQuery "Qry_kvotefiler_Oppdatere_kvitefiler_fra_zalaris", acViewNormal, acEdit
    DoCmd.OpenQuery "Qry_Oppdatere_tabell_med_korrigerte_kvoter", acViewNormal, acEdit
    MsgBox ("Kvotene er nå ferdig oppdatert!")
End If

DoCmd.SetWarnings True

End Sub

Sub test()

   DoCmd.OpenQuery "Qry_kvotefiler_finne_alle_inactive_og_withdrawn", acViewNormal, acEdit
    DoCmd.OpenQuery "Qry_kvotefiler_legge_inn_c_på_de_som_mangler_kvotekode", acViewNormal, acEdit
    DoCmd.OpenQuery "Qry_kvotefiler_korrigere_ansattdato", acViewNormal, acEdit
    DoCmd.OpenQuery "Qry_kvotefiler_oppdatere_kvotekoder", acViewNormal, acEdit
    DoCmd.OpenQuery "Qry_kvotefiler_Oppdatere_kvitefiler_fra_zalaris", acViewNormal, acEdit
    DoCmd.OpenQuery "Qry_Oppdatere_tabell_med_korrigerte_kvoter", acViewNormal, acEdit
    MsgBox ("Kvotene er nå ferdig oppdatert!")
End Sub
