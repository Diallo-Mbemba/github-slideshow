Option Strict On
Option Explicit On

''' <summary>
''' Rapport d'activité du réseau de sous-agents.
'''
''' Tout le comportement est celui de FrmRapportActivite : cette classe ne fait que poser la
''' population avant l'ouverture. Deux fenêtres, un seul code — dupliquer huit cents lignes
''' pour changer un filtre aurait garanti que les deux divergent au premier correctif.
'''
''' Les Accounts NON PARAMÉTRÉS apparaissent ici, et non chez les agences : un Account inconnu
''' n'est pas une agence propre, il n'est rien encore — et c'est dans l'écran des sous-agents
''' qu'on ira le créer.
''' </summary>
Public Class FrmRapportSousAgents
    Inherits FrmRapportActivite

    Public Sub New()
        MyBase.New()
        Portee = PorteeRapportWU.SousAgents
    End Sub

End Class
