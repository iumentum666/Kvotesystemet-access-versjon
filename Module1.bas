Attribute VB_Name = "Module1"
Option Compare Database

Public Function finne_kunde() As String

    Dim henvendelse As Long
    On Error GoTo ErrHandler:
    ' Open the form in dialog mode.
    ' This "halts" execution until the called form is closed or hidden
    DoCmd.OpenForm "Frm_Søke_på_kunder", WindowMode:=acDialog
    
    ' Since the form was opened as a dialog, we will not reach this line until it is hidden.
    ' Here, we will retrieve the value in the password text box.
    henvendelse = Forms!Frm_Søke_på_kunder.Kundenr & vbNullString
    
    ' Now, we will actually close the password form
    DoCmd.Close acForm, "Frm_Søke_på_kunder"
    
    ' And finally, we return the value we retrieved
    finne_kunde = henvendelse
ErrHandler:
End Function
