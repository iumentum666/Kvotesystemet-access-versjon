VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Frm_hovedmeny"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Command0_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
importere_fil
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Command14_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

DoCmd.SetWarnings False
DoCmd.OpenQuery "Qry_ISS_og_Arendalsansatte_oppdatere_kvotefiler_Fra_zalaris", acViewNormal, acEdit
DoCmd.OpenQuery "Qry_Oppdatere_tabell_med_korrigerte_kvoter", acViewNormal, acEdit
DoCmd.SetWarnings True

End Sub

Private Sub Command15_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
importere_salgstall
DoCmd.RunMacro "Mcr_salgstall_rydde_opp_i_salgstall"
DoCmd.OpenQuery "Qry_salgstall_finne_kunder_som_ikke_ligger_i_kundefil", acViewNormal, acEdit
DoCmd.OpenQuery "Qry_salgstall_finne_materialer_som_mangler_masterdata", acViewNormal, acEdit
End Sub

Private Sub Command6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
importere_salgstall
End Sub

Private Sub Command7_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
DoCmd.OpenQuery "Qry_salgstall_finne_materialer_som_mangler_masterdata", acViewNormal, acEdit
End Sub

