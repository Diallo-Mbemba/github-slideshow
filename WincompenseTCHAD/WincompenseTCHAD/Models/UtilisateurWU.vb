Option Strict On
Option Explicit On

''' <summary>Rôles prévus par la banque.</summary>
Public Enum RoleWU
    ''' <summary>Rôle non reconnu : aucun droit.</summary>
    Inconnu = 0
    ''' <summary>Agent de la compense : charge les rapports, calcule, génère les pièces comptables.</summary>
    Compense = 1
    ''' <summary>Commercial : crée et modifie les sous-agents, les agences et les groupes.</summary>
    Commercial = 2
    ''' <summary>Administrateur : les deux, plus les comptes systèmes et les utilisateurs.</summary>
    Administrateur = 3
End Enum

''' <summary>
''' Un utilisateur de l'application (table T_UtilisateurWU).
'''
''' Le mot de passe n'apparaît nulle part : seule son empreinte PBKDF2 est conservée, avec son
''' sel et son nombre d'itérations. Conserver ce nombre en base permet de le relever plus tard
''' sans invalider les comptes existants — chaque empreinte se vérifie avec le sien.
''' </summary>
Public Class UtilisateurWU

#Region "Libellés des rôles, tels qu'enregistrés"

    Public Const ROLE_COMPENSE As String = "COMPENSE"
    Public Const ROLE_COMMERCIAL As String = "COMMERCIAL"
    Public Const ROLE_ADMIN As String = "ADMIN"

#End Region

    Public Property Identifiant As String = String.Empty
    Public Property NomComplet As String = String.Empty
    Public Property Role As RoleWU = RoleWU.Inconnu

    Public Property Empreinte As String = String.Empty
    Public Property Sel As String = String.Empty
    Public Property Iterations As Integer = 0

    Public Property Actif As Boolean = True
    Public Property DoitChangerMotDePasse As Boolean = False

    Public Property EchecsConsecutifs As Integer = 0
    Public Property DateVerrouillage As Date?
    Public Property DerniereConnexion As Date?

    Public Property DateCreation As Date?
    Public Property CreePar As String = String.Empty
    Public Property DateModification As Date?
    Public Property ModifiePar As String = String.Empty

#Region "Droits"

    ''' <summary>
    ''' Charger les rapports, calculer, générer les pièces comptables et les historiser.
    ''' </summary>
    Public ReadOnly Property PeutTraiterLaCompense As Boolean
        Get
            Return Role = RoleWU.Compense OrElse Role = RoleWU.Administrateur
        End Get
    End Property

    ''' <summary>
    ''' Créer et modifier les sous-agents, les agences propres et les groupes statistiques.
    ''' L'agent de compense n'y a AUCUN accès, pas même en lecture : décision de la banque.
    ''' </summary>
    Public ReadOnly Property PeutGererLesPointsDeVente As Boolean
        Get
            Return Role = RoleWU.Commercial OrElse Role = RoleWU.Administrateur
        End Get
    End Property

    ''' <summary>Modifier les comptes comptables de la pièce (table SystemeWU).</summary>
    Public ReadOnly Property PeutGererLesComptesSystemes As Boolean
        Get
            Return Role = RoleWU.Administrateur
        End Get
    End Property

    ''' <summary>Créer, modifier et désactiver les comptes utilisateurs.</summary>
    Public ReadOnly Property PeutGererLesUtilisateurs As Boolean
        Get
            Return Role = RoleWU.Administrateur
        End Get
    End Property

    ''' <summary>Consulter les rapports d'activité : les trois rôles y ont accès.</summary>
    Public ReadOnly Property PeutVoirLesRapports As Boolean
        Get
            Return Role <> RoleWU.Inconnu
        End Get
    End Property

#End Region

#Region "Conversion du rôle"

    ''' <summary>Rôle correspondant au libellé enregistré. Un libellé inconnu ne donne aucun droit.</summary>
    Public Shared Function RoleDepuisLibelle(libelle As String) As RoleWU

        Select Case If(libelle, String.Empty).Trim().ToUpperInvariant()
            Case ROLE_COMPENSE : Return RoleWU.Compense
            Case ROLE_COMMERCIAL : Return RoleWU.Commercial
            Case ROLE_ADMIN : Return RoleWU.Administrateur
            Case Else : Return RoleWU.Inconnu
        End Select
    End Function

    ''' <summary>Libellé enregistré correspondant au rôle.</summary>
    Public Shared Function LibelleDepuisRole(role As RoleWU) As String

        Select Case role
            Case RoleWU.Compense : Return ROLE_COMPENSE
            Case RoleWU.Commercial : Return ROLE_COMMERCIAL
            Case RoleWU.Administrateur : Return ROLE_ADMIN
            Case Else : Return String.Empty
        End Select
    End Function

    ''' <summary>Libellé lisible d'un rôle, pour l'affichage et les listes déroulantes.</summary>
    Public Shared Function LibelleLisibleDepuisRole(role As RoleWU) As String

        Select Case role
            Case RoleWU.Compense : Return "Agent de la compense"
            Case RoleWU.Commercial : Return "Commercial"
            Case RoleWU.Administrateur : Return "Administrateur"
            Case Else : Return "Rôle inconnu"
        End Select
    End Function

    ''' <summary>Libellé lisible du rôle de cet utilisateur.</summary>
    Public ReadOnly Property LibelleRole As String
        Get
            Return LibelleLisibleDepuisRole(Role)
        End Get
    End Property

    ''' <summary>Libellés lisibles des trois rôles, dans l'ordre où ils sont proposés.</summary>
    Public Shared ReadOnly Property RolesProposes As String()
        Get
            Return New String() {"Agent de la compense", "Commercial", "Administrateur"}
        End Get
    End Property

    ''' <summary>Rôle correspondant à un libellé lisible (celui d'une liste déroulante).</summary>
    Public Shared Function RoleDepuisLibelleLisible(libelle As String) As RoleWU

        Select Case If(libelle, String.Empty).Trim()
            Case "Agent de la compense" : Return RoleWU.Compense
            Case "Commercial" : Return RoleWU.Commercial
            Case "Administrateur" : Return RoleWU.Administrateur
            Case Else : Return RoleWU.Inconnu
        End Select
    End Function

#End Region

End Class
