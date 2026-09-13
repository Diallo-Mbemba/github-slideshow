Option Strict On
Option Explicit On

''' <summary>Objet du référentiel sur lequel porte une demande.</summary>
Public Enum TypeObjetWU
    Inconnu = 0
    SousAgent = 1
    Agence = 2
    Groupe = 3
End Enum

''' <summary>Nature de l'écriture demandée.</summary>
Public Enum OperationWU
    Inconnue = 0
    Creation = 1
    Modification = 2
    Suppression = 3
    ''' <summary>Report des valeurs d'un groupe sur tous les sous-agents qui le portent.</summary>
    Synchronisation = 4
End Enum

''' <summary>État d'une demande.</summary>
Public Enum StatutDemandeWU
    Inconnu = 0
    EnAttente = 1
    Autorisee = 2
    Rejetee = 3
End Enum

''' <summary>
''' Une écriture proposée sur le référentiel des points de vente, en attente d'autorisation.
'''
''' Une demande n'est pas la donnée : c'est une intention. Tant qu'elle n'est pas autorisée,
''' T_Pdv_SA, T_Pdv_EC et T_GroupeStatistique l'ignorent, et la comptabilisation quotidienne
''' aussi.
'''
''' Les trois objets du référentiel partagent presque tous leurs champs — une clé, une
''' désignation, un groupe, un taux, deux comptes, un code de rattachement. Ils tiennent donc
''' dans une seule classe, dont les propriétés portent le nom du concept et non celui de la
''' colonne de telle ou telle table. La correspondance est documentée dans
''' Scripts\09_Demandes.sql et rendue par les vues V_Demande_*.
''' </summary>
Public Class DemandeWU

#Region "Libellés, tels qu'enregistrés"

    Public Const OBJET_SOUS_AGENT As String = "SOUS_AGENT"
    Public Const OBJET_AGENCE As String = "AGENCE"
    Public Const OBJET_GROUPE As String = "GROUPE"

    Public Const OPERATION_CREATION As String = "CREATION"
    Public Const OPERATION_MODIFICATION As String = "MODIFICATION"
    Public Const OPERATION_SUPPRESSION As String = "SUPPRESSION"
    Public Const OPERATION_SYNCHRONISATION As String = "SYNCHRONISATION"

    Public Const STATUT_EN_ATTENTE As String = "EN_ATTENTE"
    Public Const STATUT_AUTORISE As String = "AUTORISE"
    Public Const STATUT_REJETE As String = "REJETE"

#End Region

    Public Property IdDemande As Long = 0L
    Public Property TypeObjet As TypeObjetWU = TypeObjetWU.Inconnu
    Public Property Operation As OperationWU = OperationWU.Inconnue
    Public Property Statut As StatutDemandeWU = StatutDemandeWU.EnAttente

#Region "Valeurs proposées"

    ''' <summary>Account du sous-agent, Account de l'agence, ou nom du groupe.</summary>
    Public Property Cle As String = String.Empty

    Public Property Designation As String = String.Empty
    Public Property GroupeStatistique As String = String.Empty
    Public Property CompteActivite As String = String.Empty
    Public Property CompteCommission As String = String.Empty
    Public Property Taux As Decimal = 0D

    ''' <summary>codeagence pour un sous-agent, CodeAgenc-Voyager pour une agence.</summary>
    Public Property CodeRattachement As String = String.Empty

#End Region

#Region "Qui a saisi, qui a décidé"

    Public Property SaisiPar As String = String.Empty
    Public Property DateSaisie As Date?
    Public Property DecidePar As String = String.Empty
    Public Property DateDecision As Date?
    Public Property MotifRejet As String = String.Empty

#End Region

#Region "Conversion depuis les objets du référentiel"

    ''' <summary>Demande portant sur un sous-agent.</summary>
    Public Shared Function DepuisSousAgent(pdv As PointDeVenteSA, operation As OperationWU) As DemandeWU

        If pdv Is Nothing Then Return Nothing

        Return New DemandeWU() With {
            .TypeObjet = TypeObjetWU.SousAgent,
            .Operation = operation,
            .Cle = pdv.CodePdv,
            .Designation = pdv.Designation,
            .GroupeStatistique = pdv.GroupeStatistique,
            .CompteActivite = pdv.CompteCompense,
            .CompteCommission = pdv.CompteCommission,
            .Taux = pdv.Taux,
            .CodeRattachement = pdv.CodeAgence
        }
    End Function

    ''' <summary>Demande portant sur une agence propre.</summary>
    Public Shared Function DepuisAgence(pdv As PointDeVenteEC, operation As OperationWU) As DemandeWU

        If pdv Is Nothing Then Return Nothing

        Return New DemandeWU() With {
            .TypeObjet = TypeObjetWU.Agence,
            .Operation = operation,
            .Cle = pdv.CodeSite,
            .Designation = pdv.Designation,
            .CodeRattachement = pdv.CodeAgenceVoyager
        }
    End Function

    ''' <summary>Demande portant sur un groupe statistique.</summary>
    Public Shared Function DepuisGroupe(groupe As GroupeStatistiqueWU, operation As OperationWU) As DemandeWU

        If groupe Is Nothing Then Return Nothing

        Return New DemandeWU() With {
            .TypeObjet = TypeObjetWU.Groupe,
            .Operation = operation,
            .Cle = groupe.Nom,
            .CompteActivite = groupe.CompteActivite,
            .CompteCommission = groupe.CompteCommission,
            .Taux = groupe.Taux
        }
    End Function

#End Region

#Region "Conversion vers les objets du référentiel"

    ''' <summary>Sous-agent décrit par cette demande.</summary>
    Public Function VersSousAgent() As PointDeVenteSA

        Return New PointDeVenteSA() With {
            .CodePdv = Cle,
            .Designation = Designation,
            .GroupeStatistique = GroupeStatistique,
            .CompteCompense = CompteActivite,
            .CompteCommission = CompteCommission,
            .Taux = Taux,
            .CodeAgence = CodeRattachement
        }
    End Function

    ''' <summary>Agence propre décrite par cette demande.</summary>
    Public Function VersAgence() As PointDeVenteEC

        Return New PointDeVenteEC() With {
            .CodeSite = Cle,
            .Designation = Designation,
            .CodeAgenceVoyager = CodeRattachement
        }
    End Function

    ''' <summary>Groupe statistique décrit par cette demande.</summary>
    Public Function VersGroupe() As GroupeStatistiqueWU

        Return New GroupeStatistiqueWU() With {
            .Nom = Cle,
            .CompteActivite = CompteActivite,
            .CompteCommission = CompteCommission,
            .Taux = Taux
        }
    End Function

#End Region

#Region "Libellés lisibles"

    Public ReadOnly Property LibelleObjet As String
        Get
            Select Case TypeObjet
                Case TypeObjetWU.SousAgent : Return "Sous-agent"
                Case TypeObjetWU.Agence : Return "Agence propre"
                Case TypeObjetWU.Groupe : Return "Groupe statistique"
                Case Else : Return "Objet inconnu"
            End Select
        End Get
    End Property

    Public ReadOnly Property LibelleOperation As String
        Get
            Select Case Operation
                Case OperationWU.Creation : Return "Création"
                Case OperationWU.Modification : Return "Modification"
                Case OperationWU.Suppression : Return "Suppression"
                Case OperationWU.Synchronisation : Return "Synchronisation des sous-agents"
                Case Else : Return "Opération inconnue"
            End Select
        End Get
    End Property

    Public ReadOnly Property LibelleStatut As String
        Get
            Select Case Statut
                Case StatutDemandeWU.EnAttente : Return "En attente"
                Case StatutDemandeWU.Autorisee : Return "Autorisée"
                Case StatutDemandeWU.Rejetee : Return "Rejetée"
                Case Else : Return "État inconnu"
            End Select
        End Get
    End Property

    ''' <summary>
    ''' Phrase résumant la demande, pour une confirmation ou un journal.
    ''' Elle ne s'appelle pas « Resume » : c'est un mot-clé de Visual Basic.
    ''' </summary>
    Public ReadOnly Property Intitule As String
        Get
            Return $"{LibelleOperation} — {LibelleObjet} « {Cle} »"
        End Get
    End Property

#End Region

#Region "Conversion des libellés enregistrés"

    Public Shared Function ObjetDepuisLibelle(libelle As String) As TypeObjetWU
        Select Case If(libelle, String.Empty).Trim().ToUpperInvariant()
            Case OBJET_SOUS_AGENT : Return TypeObjetWU.SousAgent
            Case OBJET_AGENCE : Return TypeObjetWU.Agence
            Case OBJET_GROUPE : Return TypeObjetWU.Groupe
            Case Else : Return TypeObjetWU.Inconnu
        End Select
    End Function

    Public Shared Function LibelleDepuisObjet(objet As TypeObjetWU) As String
        Select Case objet
            Case TypeObjetWU.SousAgent : Return OBJET_SOUS_AGENT
            Case TypeObjetWU.Agence : Return OBJET_AGENCE
            Case TypeObjetWU.Groupe : Return OBJET_GROUPE
            Case Else : Return String.Empty
        End Select
    End Function

    Public Shared Function OperationDepuisLibelle(libelle As String) As OperationWU
        Select Case If(libelle, String.Empty).Trim().ToUpperInvariant()
            Case OPERATION_CREATION : Return OperationWU.Creation
            Case OPERATION_MODIFICATION : Return OperationWU.Modification
            Case OPERATION_SUPPRESSION : Return OperationWU.Suppression
            Case OPERATION_SYNCHRONISATION : Return OperationWU.Synchronisation
            Case Else : Return OperationWU.Inconnue
        End Select
    End Function

    Public Shared Function LibelleDepuisOperation(operation As OperationWU) As String
        Select Case operation
            Case OperationWU.Creation : Return OPERATION_CREATION
            Case OperationWU.Modification : Return OPERATION_MODIFICATION
            Case OperationWU.Suppression : Return OPERATION_SUPPRESSION
            Case OperationWU.Synchronisation : Return OPERATION_SYNCHRONISATION
            Case Else : Return String.Empty
        End Select
    End Function

    Public Shared Function StatutDepuisLibelle(libelle As String) As StatutDemandeWU
        Select Case If(libelle, String.Empty).Trim().ToUpperInvariant()
            Case STATUT_EN_ATTENTE : Return StatutDemandeWU.EnAttente
            Case STATUT_AUTORISE : Return StatutDemandeWU.Autorisee
            Case STATUT_REJETE : Return StatutDemandeWU.Rejetee
            Case Else : Return StatutDemandeWU.Inconnu
        End Select
    End Function

    Public Shared Function LibelleDepuisStatut(statut As StatutDemandeWU) As String
        Select Case statut
            Case StatutDemandeWU.EnAttente : Return STATUT_EN_ATTENTE
            Case StatutDemandeWU.Autorisee : Return STATUT_AUTORISE
            Case StatutDemandeWU.Rejetee : Return STATUT_REJETE
            Case Else : Return String.Empty
        End Select
    End Function

#End Region

End Class
