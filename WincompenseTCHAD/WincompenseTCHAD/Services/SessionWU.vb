Option Strict On
Option Explicit On

''' <summary>
''' Utilisateur connecté pour la durée de l'exécution.
'''
''' Point d'accès unique aux droits : aucun écran ne teste un rôle directement, tous
''' interrogent cette classe. Une règle d'accès se change ainsi à un seul endroit.
'''
''' Tant que personne n'est connecté, Utilisateur vaut Nothing et TOUS les droits sont refusés :
''' une erreur d'enchaînement ne peut donc pas ouvrir l'application sans identification.
''' </summary>
Public NotInheritable Class SessionWU

    Private Sub New()
    End Sub

    ''' <summary>Utilisateur connecté, ou Nothing.</summary>
    Public Shared Property Utilisateur As UtilisateurWU

    ''' <summary>Identifiant de l'utilisateur connecté, ou chaîne vide. Sert à la traçabilité.</summary>
    Public Shared ReadOnly Property Identifiant As String
        Get
            Return If(Utilisateur Is Nothing, String.Empty, Utilisateur.Identifiant)
        End Get
    End Property

    ''' <summary>Nom complet et rôle, pour la barre d'état.</summary>
    Public Shared ReadOnly Property Description As String
        Get
            If Utilisateur Is Nothing Then Return "Non connecté"
            Return $"{Utilisateur.NomComplet} ({Utilisateur.Identifiant}) — {Utilisateur.LibelleRole}"
        End Get
    End Property

    ''' <summary>
    ''' Identifiant à inscrire dans les colonnes de traçabilité.
    '''
    ''' Personne n'est censé écrire sans être connecté, l'application exigeant une
    ''' identification avant d'afficher le moindre écran. La valeur de repli existe pour que
    ''' l'écriture aboutisse tout de même — une trace incomplète vaut mieux qu'une écriture
    ''' perdue — et pour que l'anomalie se voie dans la colonne plutôt que de passer inaperçue.
    ''' </summary>
    Public Shared ReadOnly Property Auteur As String
        Get
            ' La variable ne s'appelle pas « identifiant » : VB ne distingue pas la casse, elle
            ' masquerait la propriété Identifiant qu'elle est censée lire.
            Dim courant As String = Identifiant
            Return If(String.IsNullOrEmpty(courant), "non identifié", courant)
        End Get
    End Property

    ''' <summary>Nom de la machine, pour le journal des connexions.</summary>
    Public Shared ReadOnly Property Poste As String
        Get
            Try
                Return Environment.MachineName
            Catch ex As InvalidOperationException
                Return String.Empty
            End Try
        End Get
    End Property

    ''' <summary>
    ''' Session Windows sous laquelle l'application tourne. Les postes étant nominatifs, elle
    ''' permet de rapprocher une connexion applicative de la session du poste — et de repérer
    ''' un identifiant utilisé depuis un poste qui n'est pas le sien.
    ''' </summary>
    Public Shared ReadOnly Property CompteWindows As String
        Get
            Try
                Return Environment.UserName
            Catch ex As InvalidOperationException
                Return String.Empty
            End Try
        End Get
    End Property

#Region "Droits"

    Public Shared ReadOnly Property PeutTraiterLaCompense As Boolean
        Get
            Return Utilisateur IsNot Nothing AndAlso Utilisateur.PeutTraiterLaCompense
        End Get
    End Property

    Public Shared ReadOnly Property PeutGererLesPointsDeVente As Boolean
        Get
            Return Utilisateur IsNot Nothing AndAlso Utilisateur.PeutGererLesPointsDeVente
        End Get
    End Property

    Public Shared ReadOnly Property PeutGererLesComptesSystemes As Boolean
        Get
            Return Utilisateur IsNot Nothing AndAlso Utilisateur.PeutGererLesComptesSystemes
        End Get
    End Property

    Public Shared ReadOnly Property PeutGererLesUtilisateurs As Boolean
        Get
            Return Utilisateur IsNot Nothing AndAlso Utilisateur.PeutGererLesUtilisateurs
        End Get
    End Property

    Public Shared ReadOnly Property PeutVoirLesRapports As Boolean
        Get
            Return Utilisateur IsNot Nothing AndAlso Utilisateur.PeutVoirLesRapports
        End Get
    End Property

#End Region

    ''' <summary>Ouvre la session.</summary>
    Public Shared Sub Ouvrir(utilisateurConnecte As UtilisateurWU)
        Utilisateur = utilisateurConnecte
    End Sub

    ''' <summary>Ferme la session : tous les droits retombent immédiatement.</summary>
    Public Shared Sub Fermer()
        Utilisateur = Nothing
    End Sub

End Class
