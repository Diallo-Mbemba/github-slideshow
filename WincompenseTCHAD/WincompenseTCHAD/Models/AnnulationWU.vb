Option Strict On
Option Explicit On

''' <summary>Raison pour laquelle une journée comptabilisée doit être retirée.</summary>
Public Enum MotifAnnulationWU
    Inconnu = 0
    ''' <summary>Le rapport Western Union ne contenait aucune opération.</summary>
    RapportVide = 1
    ''' <summary>Le rapport contenait des données fausses.</summary>
    RapportErrone = 2
    ''' <summary>Le rapport chargé ne portait pas sur cette journée.</summary>
    MauvaiseJournee = 3
    ''' <summary>La journée a été comptabilisée deux fois.</summary>
    Doublon = 4
    Autre = 9
End Enum

''' <summary>
''' Ce qu'une annulation de comptabilisation demande, et ce qu'elle laisse derrière elle.
'''
''' ANNULER N'EST PAS SUPPRIMER. Les lignes de la journée sont déplacées vers les tables
''' d'archive, avec le motif et les deux signatures : la journée sort des rapports, elle ne
''' sort pas de la mémoire de l'application. C'est précisément quand une journée est annulée
''' que l'auditeur veut savoir qui, quand et pourquoi.
'''
''' ET SURTOUT : ANNULER N'EXTOURNE RIEN. Si le fichier destiné au core banking a déjà été
''' injecté, les écritures sont dans les livres de la banque et Wincompense ne peut pas les
''' en retirer. La réponse de l'agent sur ce point est donc conservée, et l'écran le dit en
''' toutes lettres avant de déposer la demande.
''' </summary>
Public Class AnnulationWU

#Region "Motifs, tels qu'enregistrés"

    Public Const MOTIF_RAPPORT_VIDE As String = "RAPPORT_VIDE"
    Public Const MOTIF_RAPPORT_ERRONE As String = "RAPPORT_ERRONE"
    Public Const MOTIF_MAUVAISE_JOURNEE As String = "MAUVAISE_JOURNEE"
    Public Const MOTIF_DOUBLON As String = "DOUBLON"
    Public Const MOTIF_AUTRE As String = "AUTRE"

#End Region

#Region "La demande"

    Public Property DateActivite As Date
    Public Property Motif As MotifAnnulationWU = MotifAnnulationWU.Inconnu
    Public Property Commentaire As String = String.Empty

    ''' <summary>
    ''' Réponse de l'agent à la question « le fichier core banking a-t-il déjà été injecté ? ».
    ''' L'application ne peut pas la déduire : elle sait qu'elle a produit le fichier, pas ce
    ''' que le core banking en a fait.
    ''' </summary>
    Public Property CoreBankingInjecte As Boolean = False

#End Region

#Region "Ce qui a été retiré, et par qui"

    Public Property IdAnnulation As Long = 0L
    Public Property NombrePdv As Integer = 0
    Public Property NombreTransactions As Integer = 0
    Public Property NombreEcritures As Integer = 0
    Public Property TotalDebit As Long = 0L
    Public Property TotalCredit As Long = 0L

    Public Property DemandeePar As String = String.Empty
    Public Property DateDemande As Date?
    Public Property AutoriseePar As String = String.Empty
    Public Property DateAutorisation As Date?

#End Region

#Region "Libellés lisibles"

    Public ReadOnly Property LibelleMotif As String
        Get
            Return LibelleLisible(Motif)
        End Get
    End Property

    ''' <summary>Phrase résumant l'annulation, pour une confirmation ou un journal.</summary>
    Public ReadOnly Property Intitule As String
        Get
            Return $"Annulation de la comptabilisation du {DateActivite:dd/MM/yyyy} — {LibelleMotif}"
        End Get
    End Property

#End Region

#Region "Conversion des libellés enregistrés"

    Public Shared Function MotifDepuisCode(code As String) As MotifAnnulationWU

        Select Case If(code, String.Empty).Trim().ToUpperInvariant()
            Case MOTIF_RAPPORT_VIDE : Return MotifAnnulationWU.RapportVide
            Case MOTIF_RAPPORT_ERRONE : Return MotifAnnulationWU.RapportErrone
            Case MOTIF_MAUVAISE_JOURNEE : Return MotifAnnulationWU.MauvaiseJournee
            Case MOTIF_DOUBLON : Return MotifAnnulationWU.Doublon
            Case MOTIF_AUTRE : Return MotifAnnulationWU.Autre
            Case Else : Return MotifAnnulationWU.Inconnu
        End Select
    End Function

    Public Shared Function CodeDepuisMotif(valeur As MotifAnnulationWU) As String

        Select Case valeur
            Case MotifAnnulationWU.RapportVide : Return MOTIF_RAPPORT_VIDE
            Case MotifAnnulationWU.RapportErrone : Return MOTIF_RAPPORT_ERRONE
            Case MotifAnnulationWU.MauvaiseJournee : Return MOTIF_MAUVAISE_JOURNEE
            Case MotifAnnulationWU.Doublon : Return MOTIF_DOUBLON
            Case MotifAnnulationWU.Autre : Return MOTIF_AUTRE
            Case Else : Return String.Empty
        End Select
    End Function

    ''' <summary>Le motif en clair, tel qu'il s'affiche à l'écran et sur l'archive.</summary>
    Public Shared Function LibelleLisible(valeur As MotifAnnulationWU) As String

        Select Case valeur
            Case MotifAnnulationWU.RapportVide : Return "Rapport Western Union sans aucune opération"
            Case MotifAnnulationWU.RapportErrone : Return "Rapport Western Union erroné"
            Case MotifAnnulationWU.MauvaiseJournee : Return "Rapport d'une autre journée"
            Case MotifAnnulationWU.Doublon : Return "Journée comptabilisée deux fois"
            Case MotifAnnulationWU.Autre : Return "Autre motif"
            Case Else : Return "Motif non précisé"
        End Select
    End Function

    ''' <summary>
    ''' Le motif en clair à partir du code enregistré, pour l'affichage d'une archive.
    ''' Un code inconnu est rendu tel quel plutôt que traduit en « non précisé » : mieux vaut
    ''' montrer une valeur qu'on ne comprend pas que faire disparaître ce qui est écrit.
    ''' </summary>
    Public Shared Function LibelleDepuisCode(code As String) As String

        Dim valeur As MotifAnnulationWU = MotifDepuisCode(code)
        If valeur = MotifAnnulationWU.Inconnu Then Return If(code, String.Empty).Trim()
        Return LibelleLisible(valeur)
    End Function

#End Region

#Region "Choix proposé à l'écran"

    ''' <summary>Un motif tel qu'il apparaît dans une liste déroulante.</summary>
    Public NotInheritable Class ChoixMotif

        Public ReadOnly Property Valeur As MotifAnnulationWU
        Public ReadOnly Property Libelle As String

        Public Sub New(motifChoisi As MotifAnnulationWU)
            _Valeur = motifChoisi
            _Libelle = LibelleLisible(motifChoisi)
        End Sub

        Public Overrides Function ToString() As String
            Return Libelle
        End Function
    End Class

    ''' <summary>
    ''' Les motifs proposés, dans l'ordre où ils se rencontrent. « Autre » ferme la liste :
    ''' il oblige à écrire un commentaire, et n'est donc pas le choix le plus rapide.
    ''' </summary>
    Public Shared Function MotifsProposes() As List(Of ChoixMotif)

        Return New List(Of ChoixMotif) From {
            New ChoixMotif(MotifAnnulationWU.RapportVide),
            New ChoixMotif(MotifAnnulationWU.RapportErrone),
            New ChoixMotif(MotifAnnulationWU.MauvaiseJournee),
            New ChoixMotif(MotifAnnulationWU.Doublon),
            New ChoixMotif(MotifAnnulationWU.Autre)
        }
    End Function

#End Region

End Class
