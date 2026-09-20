Option Strict On
Option Explicit On

''' <summary>
''' Rapport d'activité du réseau propre de la banque.
'''
''' Même code que le rapport des sous-agents, autre population — et donc d'autres axes :
''' l'onglet « par groupe statistique » cède la place à la PERFORMANCE DES AGENCES, et le
''' filtre porte sur l'agence au lieu du groupe.
'''
''' Ce n'est pas une préférence de présentation : T_Pdv_EC n'a pas de colonne groupe, une
''' agence propre ne porte aucun taux de rétrocession, et sa commission ne se partage pas.
''' Les trois axes du rapport des sous-agents n'ont ici aucun objet.
''' </summary>
Public Class FrmRapportAgences
    Inherits FrmRapportActivite

    Public Sub New()
        MyBase.New()
        Portee = PorteeRapportWU.AgencesPropres
    End Sub

End Class
