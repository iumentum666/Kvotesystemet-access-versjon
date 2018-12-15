Attribute VB_Name = "Mdl_oppdatering_av_salgstall"
Option Compare Database

Sub importere_salgstall()

MsgBox ("Velkommen til filoppdatereren. Velg filen med salgstall i neste dialogboks.")

Dim fd As FileDialog
Dim wb1 As Excel.Workbook
Dim rader As Integer

Set fd = Application.FileDialog(msoFileDialogFilePicker)

'get the number of the button chosen
Dim FileChosen As Integer
FileChosen = fd.Show

'Sletter alt innhold i tabell
DoCmd.SetWarnings False
DoCmd.RunSQL "DELETE * FROM Tbl_salgstall_for_import"
DoCmd.SetWarnings True

If FileChosen <> -1 Then
    'Melding om at man trykket cancel
    MsgBox "Du valgte cancel"
Else
    'Her gjør vi alle endringene i databasen
    Set wb1 = Workbooks.Open(fd.SelectedItems(1))
    wb1.Sheets("Sheet1").Cells(1, 12) = "Field12"
    wb1.Sheets("Sheet1").Cells(1, 14) = "Field14"
    wb1.Sheets("Sheet1").Cells(1, 15) = "Field15"
    wb1.Sheets("Sheet1").Cells(1, 36) = "Field33"
    wb1.Sheets("Sheet1").Cells(1, 46) = "Field43"
    wb1.Sheets("Sheet1").Cells(1, 10) = "Custpricprocedure"
    wb1.Sheets("Sheet1").Cells(1, 17) = "Sales Representative no"
    wb1.Sheets("Sheet1").Cells(1, 18) = "Sales Rep Name"
    wb1.Sheets("Sheet1").Cells(1, 19) = "Requested delivdate"
    wb1.Sheets("Sheet1").Cells(1, 26) = "Customer classific"
    wb1.Sheets("Sheet1").Cells(1, 30) = "Purchase order no"
    rader = wb1.Sheets("Sheet1").UsedRange.Rows.Count
    If rader > 1500 Then
        wb1.Sheets.Copy after:=Sheets(Sheets.Count)
        ActiveSheet.Name = "Sheet2"
        wb1.Sheets("Sheet1").Activate
            Rows("2:1499").Select
            Selection.Delete
        wb1.Sheets("Sheet2").Activate
            Rows("1500:" & rader).Select
            Selection.Delete
    End If
    wb1.Close (True)
    'DoCmd.TransferSpreadsheet acImport, , "Tbl_salgstall_for_import", _
    fd.SelectedItems(1), True, "Sheet1"
    DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12Xml, "Tbl_salgstall_for_import", fd.SelectedItems(1), True, "Sheet1!" & fullranges
    If rader > 1500 Then
        DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12Xml, "Tbl_salgstall_for_import", fd.SelectedItems(1), True, "Sheet2!" & fullranges
    End If
    MsgBox ("Salgstallene er importert!")
End If

End Sub
