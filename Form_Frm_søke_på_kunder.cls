VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Frm_søke_på_kunder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Command56_Click()
Me.Visible = False
End Sub

Private Sub Text13_AfterUpdate()
soke_kunde
End Sub
Private Sub Text15_AfterUpdate()
soke_kunde
End Sub
Private Sub Text47_AfterUpdate()
soke_kunde
End Sub

Private Sub Text49_AfterUpdate()
soke_kunde
End Sub

Private Sub Text51_AfterUpdate()
soke_kunde
End Sub

Private Sub Text53_AfterUpdate()
soke_kunde
End Sub
Private Sub Text63_AfterUpdate()
soke_kunde
End Sub

Sub soke_kunde()

Dim filterstreng As String
Dim ferdig_filter As String

Me.FilterOn = False

If Not Text13.Value = "" Then
    filterstreng = filterstreng & "[Ansattnummer Field] " & " Like ""*" & Text13.Value & "*""" & " And "
End If

If Not Text15.Value = "" Then
    filterstreng = filterstreng & "[Fornavn Field] " & "LIKE ""*" & Text15.Value & "*""" & " AND "
End If

If Not Text47.Value = "" Then
    filterstreng = filterstreng & "[Etternavn Field] " & "LIKE ""*" & Text47.Value & "*""" & " AND "
End If

If Not Text49.Value = "" Then
    filterstreng = filterstreng & "[C47 nummer] " & "LIKE ""*" & Text49.Value & "*""" & " AND "
End If

If Not Text63.Value = "" Then
    filterstreng = filterstreng & "[T0E] " & "LIKE ""*" & Text63.Value & "*""" & " AND "
End If

If Not Len(filterstreng) = 0 Then
    ferdig_filter = Left(filterstreng, Len(filterstreng) - 5)
    Me.Filter = ferdig_filter
    Me.FilterOn = True
End If

End Sub





