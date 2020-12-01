Attribute VB_Name = "Cours1"
'Mon premier module
Option Explicit


'Calculez le salaire hebdomadaire
Sub salaire_hebdomadaire()

Dim salaire_horaire As Double
Dim nombre_heures_travaillees As Double
Dim salaire_hebdomadaire As Double

'Assignation du salaire et nombre d'heures
salaire_horaire = 219.12
nombre_heures_travaillees = 12

'Calcul du salaire hebdomadaire
salaire_hebdomadaire = salaire_horaire * nombre_heures_travaillees

'Affichage du salaire
Call MsgBox(salaire_hebdomadaire)

End Sub

'Calculez le salaire hebdomadaire
Sub salaire_hebdomadaire_avec_inputbox()

Dim salaire_horaire As Double
Dim nombre_heures_travaillees As Double
Dim salaire_hebdomadaire As Double

'Assignation du salaire et nombre d'heures
salaire_horaire = Val(InputBox("Votre salaire horaire"))
nombre_heures_travaillees = Val(InputBox("Votre nombres d'heures travaill�es"))

'Calcul du salaire hebdomadaire
salaire_hebdomadaire = salaire_horaire * nombre_heures_travaillees

'Affichage du salaire
Call MsgBox(salaire_hebdomadaire)

End Sub

'Calculez le salaire hebdomadaire
Sub salaire_hebdomadaire_avec_phrase()

Dim salaire_horaire As Double
Dim nombre_heures_travaillees As Double
Dim salaire_hebdomadaire As Double

'Assignation du salaire et nombre d'heures
salaire_horaire = Val(InputBox("Votre salaire horaire"))
nombre_heures_travaillees = Val(InputBox("Votre nombres d'heures travaill�es"))

'Calcul du salaire hebdomadaire
salaire_hebdomadaire = salaire_horaire * nombre_heures_travaillees

'Affichage du salaire
Call MsgBox("Le salaire hebdomadaire est : " & salaire_hebdomadaire & "$")

End Sub

'Calculez le salaire hebdomadaire
Sub salaire_hebdomadaire_avec_temps_supplementaire()

Dim salaire_horaire As Double
Dim nombre_heures_travaillees As Double
Dim salaire_hebdomadaire As Double

'Assignation du salaire et nombre d'heures
salaire_horaire = Val(InputBox("Votre salaire horaire"))
nombre_heures_travaillees = Val(InputBox("Votre nombres d'heures travaill�es"))

'Calcul du salaire hebdomadaire
If nombre_heures_travaillees <= 40 Then
    salaire_hebdomadaire = salaire_horaire * nombre_heures_travaillees
Else
    salaire_hebdomadaire = salaire_horaire * 40 + salaire_horaire * 2 * (nombre_heures_travaillees - 40)
End If

'Affichage du salaire
Call MsgBox("Le salaire hebdomadaire est : " & salaire_hebdomadaire & "$")

End Sub


