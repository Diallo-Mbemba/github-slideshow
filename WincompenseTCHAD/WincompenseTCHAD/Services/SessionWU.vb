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
            ' La variable ne s'appelle pas « description » : VB ne distingue pas la casse, elle
            ' masquerait la propriété Description à l'intérieur de laquelle elle est déclarée.
            Dim libelle As String =
                $"{Utilisateur.NomComplet} ({Utilisateur.Identifiant}) — {Utilisateur.LibelleRole}"

            If Utilisateur.Fonction = FonctionWU.Aucune Then Return libelle

            Return $"{libelle} · {Utilisateur.LibelleFonction}"
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

    ''' <summary>Déposer une saisie sur le référentiel des points de vente.</summary>
    Public Shared ReadOnly Property PeutSaisirLesPointsDeVente As Boolean
        Get
            Return Utilisateur IsNot Nothing AndAlso Utilisateur.PeutSaisirLesPointsDeVente
        End Get
    End Property

    ''' <summary>Autoriser ou rejeter les saisies déposées par d'autres.</summary>
    Public Shared ReadOnly Property PeutAutoriserLesPointsDeVente As Boolean
        Get
            Return Utilisateur IsNot Nothing AndAlso Utilisateur.PeutAutoriserLesPointsDeVente
        End Get
    End Property

    Public Shared ReadOnly Property PeutGererLesComptesSystemes As Boolean
        Get
            Return Utilisateur IsNot Nothing AndAlso Utilisateur.PeutGererLesComptesSystemes
        End Get
    End Property

    ''' <summary>Ouvrir l'écran des utilisateurs : l'administrateur et le contrôleur.</summary>
    Public Shared ReadOnly Property PeutGererLesUtilisateurs As Boolean
        Get
            Return Utilisateur IsNot Nothing AndAlso Utilisateur.PeutGererLesUtilisateurs
        End Get
    End Property

    ''' <summary>Changer le serveur et la base : l'administrateur seul.</summary>
    Public Shared ReadOnly Property PeutChangerLaConnexion As Boolean
        Get
            Return Utilisateur IsNot Nothing AndAlso Utilisateur.PeutChangerLaConnexion
        End Get
    End Property

    ''' <summary>Modifier le nom, le rôle ou la fonction d'un compte : l'administrateur seul.</summary>
    Public Shared ReadOnly Property PeutModifierLesComptes As Boolean
        Get
            Return Utilisateur IsNot Nothing AndAlso Utilisateur.PeutModifierLesComptes
        End Get
    End Property

    ''' <summary>Réinitialiser un mot de passe : l'administrateur seul.</summary>
    Public Shared ReadOnly Property PeutReinitialiserLesMotsDePasse As Boolean
        Get
            Return Utilisateur IsNot Nothing AndAlso Utilisateur.PeutReinitialiserLesMotsDePasse
        End Get
    End Property

    ''' <summary>
    ''' Activer, désactiver ou déverrouiller CE compte-là.
    '''
    ''' Le seul droit de cette classe qui dépende d'une cible : un contrôleur n'agit sur aucun
    ''' administrateur, ni sur son propre compte. Les écrans l'interrogent ligne par ligne.
    ''' </summary>
    Public Shared Function PeutAgirSurLActivite(cible As UtilisateurWU) As Boolean
        Return Utilisateur IsNot Nothing AndAlso Utilisateur.PeutAgirSurLActivite(cible)
    End Function

    ''' <summary>Créer un compte portant ce rôle. Voir UtilisateurWU.PeutCreerLeRole.</summary>
    Public Shared Function PeutCreerLeRole(role As RoleWU) As Boolean
        Return Utilisateur IsNot Nothing AndAlso Utilisateur.PeutCreerLeRole(role)
    End Function

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
