Option Strict On
Option Explicit On

''' <summary>
''' Une transaction Western Union identifiée par son MTCN : une ligne du rapport d'activité,
''' conservée telle quelle dans l'historique détaillé (table T_HistoriqueMTCN).
'''
''' L'agrégat de la journée ne peut pas porter cette information — il regroupe justement les
''' transactions — d'où cette seconde table, qui permet de retrouver et de justifier une
''' opération précise.
''' </summary>
Public Class TransactionWU

    ''' <summary>Libellés des deux sens, tels qu'enregistrés et affichés.</summary>
    Public Const SENS_ENVOI As String = "ENVOI"
    Public Const SENS_PAIEMENT As String = "PAIEMENT"

    Public Property DateActivite As Date
    Public Property Account As String = String.Empty

    ''' <summary>Référence Western Union de la transaction.</summary>
    Public Property MTCN As String = String.Empty

    ''' <summary>ENVOI ou PAIEMENT, d'après SendPayIndicator.</summary>
    Public Property Sens As String = String.Empty

    ''' <summary>Statut Western Union : S (réglée), W (en attente), C (annulée).</summary>
    Public Property Statut As String = String.Empty

    ''' <summary>Principal envoyé pour un envoi, principal payé pour un paiement.</summary>
    Public Property Montant As Decimal = 0D

    ''' <summary>Identification du point de vente, recopiée telle qu'elle était ce jour-là.</summary>
    Public Property Designation As String = String.Empty
    Public Property GroupeStatistique As String = String.Empty
    Public Property TypePdv As String = String.Empty

    ''' <summary>Vrai si la transaction a été annulée : elle est conservée, mais hors des montants agrégés.</summary>
    Public ReadOnly Property EstAnnulee As Boolean
        Get
            For Each statutExclu As String In ConstantesWU.StatutsActiviteExclus
                If String.Equals(Statut, statutExclu, StringComparison.OrdinalIgnoreCase) Then Return True
            Next
            Return False
        End Get
    End Property

End Class
