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
    ''' <summary>
    ''' Contrôleur : consulte les rapports et administre les comptes utilisateurs, sans aucun
    ''' accès à la compense ni au référentiel.
    '''
    ''' IL N'EST PAS UN ADMINISTRATEUR AU RABAIS. Son pouvoir sur les comptes est volontairement
    ''' borné : il crée des agents et des commerciaux, il active et désactive, et c'est tout. Il
    ''' ne touche NI un administrateur, NI un mot de passe, NI le rôle d'un compte existant —
    ''' chacune de ces trois portes suffirait à lui rendre les droits qu'on ne lui donne pas.
    ''' </summary>
    Controleur = 4
End Enum

''' <summary>
''' Fonction de l'utilisateur dans le double regard sur le référentiel des points de vente.
'''
''' Le rôle dit le domaine — compense, commercial, administration. La fonction dit le pouvoir :
''' saisir, ou autoriser. Les deux ne se cumulent pas sur une même personne, faute de quoi le
''' double regard n'en serait plus un.
''' </summary>
Public Enum FonctionWU
    ''' <summary>Aucune : l'utilisateur ne peut ni saisir ni autoriser le référentiel.</summary>
    Aucune = 0
    ''' <summary>Saisit les créations, modifications et suppressions. Elles restent en attente.</summary>
    Inputer = 1
    ''' <summary>Autorise ou rejette les saisies des autres. Jamais les siennes.</summary>
    Authorizer = 2
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
    Public Const ROLE_CONTROLEUR As String = "CONTROLEUR"

    Public Const FONCTION_INPUTER As String = "INPUTER"
    Public Const FONCTION_AUTHORIZER As String = "AUTHORIZER"

#End Region

    Public Property Identifiant As String = String.Empty
    Public Property NomComplet As String = String.Empty
    Public Property Role As RoleWU = RoleWU.Inconnu

    ''' <summary>Fonction dans le double regard sur le référentiel. Aucune par défaut.</summary>
    Public Property Fonction As FonctionWU = FonctionWU.Aucune

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

    ''' <summary>
    ''' Déposer une création, une modification ou une suppression sur le référentiel.
    '''
    ''' Le droit ne suffit pas à écrire : la saisie part en attente et n'existe pour personne
    ''' tant qu'un authorizer ne l'a pas autorisée.
    ''' </summary>
    Public ReadOnly Property PeutSaisirLesPointsDeVente As Boolean
        Get
            Return PeutGererLesPointsDeVente AndAlso Fonction = FonctionWU.Inputer
        End Get
    End Property

    ''' <summary>
    ''' Autoriser ou rejeter les saisies déposées par d'autres.
    '''
    ''' Jamais les siennes : la règle ne dépend pas de ce droit mais de l'identité de celui qui
    ''' a saisi, et elle est vérifiée à la décision — y compris pour un administrateur.
    ''' </summary>
    Public ReadOnly Property PeutAutoriserLesPointsDeVente As Boolean
        Get
            Return PeutGererLesPointsDeVente AndAlso Fonction = FonctionWU.Authorizer
        End Get
    End Property

    ''' <summary>Modifier les comptes comptables de la pièce (table SystemeWU).</summary>
    Public ReadOnly Property PeutGererLesComptesSystemes As Boolean
        Get
            Return Role = RoleWU.Administrateur
        End Get
    End Property

    ''' <summary>
    ''' Ouvrir l'écran des utilisateurs et des connexions.
    '''
    ''' CE DROIT N'EST PLUS CELUI DE TOUT FAIRE. L'administrateur et le contrôleur ouvrent le
    ''' même écran, et n'y disposent pas des mêmes boutons : ce que chacun peut y faire se lit
    ''' dans les quatre droits ci-dessous, et sur la LIGNE SÉLECTIONNÉE pour l'activation.
    ''' </summary>
    Public ReadOnly Property PeutGererLesUtilisateurs As Boolean
        Get
            Return Role = RoleWU.Administrateur OrElse Role = RoleWU.Controleur
        End Get
    End Property

    ''' <summary>
    ''' Changer le serveur et la base de l'application.
    '''
    ''' SÉPARÉ DE LA GESTION DES UTILISATEURS, ET CE N'EST PAS UN RAFFINEMENT. Les deux entrées
    ''' du menu Sécurité partageaient la même condition : ouvrir l'écran des utilisateurs au
    ''' contrôleur lui aurait donné, du même geste, celui qui déplace la base de TOUTE la
    ''' banque. Un droit de plus coûte une ligne ; cette porte-là aurait coûté une journée.
    ''' </summary>
    Public ReadOnly Property PeutChangerLaConnexion As Boolean
        Get
            Return Role = RoleWU.Administrateur
        End Get
    End Property

    ''' <summary>
    ''' Modifier un compte existant : son nom, son RÔLE, sa fonction dans le double regard.
    '''
    ''' L'administrateur seul. Le contrôleur en est écarté parce que cet écran porte le rôle :
    ''' le lui ouvrir reviendrait à lui laisser promouvoir quelqu'un administrateur, puis se
    ''' connecter sous ce compte. Activer et désactiver, qui est son vrai besoin, passe par
    ''' <see cref="PeutAgirSurLActivite"/> et ne touche à rien d'autre.
    ''' </summary>
    Public ReadOnly Property PeutModifierLesComptes As Boolean
        Get
            Return Role = RoleWU.Administrateur
        End Get
    End Property

    ''' <summary>
    ''' Réinitialiser le mot de passe d'un compte.
    '''
    ''' L'administrateur seul. Réinitialiser le mot de passe d'un administrateur, c'est prendre
    ''' son compte — plus sûrement que le désactiver. Laisser ce bouton au contrôleur viderait
    ''' de son sens la règle qui lui interdit de toucher aux administrateurs.
    ''' </summary>
    Public ReadOnly Property PeutReinitialiserLesMotsDePasse As Boolean
        Get
            Return Role = RoleWU.Administrateur
        End Get
    End Property

    ''' <summary>
    ''' Activer, désactiver ou déverrouiller le compte désigné.
    '''
    ''' LA RÈGLE DU CONTRÔLEUR SE LIT ICI, ET NULLE PART AILLEURS : il agit sur tout compte qui
    ''' n'est pas administrateur, et sur aucun administrateur — celui qui l'a créé compris.
    ''' Avoir retenu « aucun administrateur » plutôt que « celui qui m'a créé » n'est pas une
    ''' sévérité gratuite : la seconde formulation se vérifie en remontant qui a créé qui, et
    ''' elle perd son objet dès qu'un contrôleur est créé par un contrôleur.
    '''
    ''' SON PROPRE COMPTE LUI EST FERMÉ. Se désactiver soi-même n'est pas un droit, c'est une
    ''' sortie sans retour : il faudrait un administrateur pour rouvrir le compte de celui qui
    ''' vient de fermer le sien.
    ''' </summary>
    ''' <param name="cible">Le compte sur lequel on veut agir. Nothing ne donne aucun droit.</param>
    Public Function PeutAgirSurLActivite(cible As UtilisateurWU) As Boolean

        If cible Is Nothing Then Return False

        Select Case Role
            Case RoleWU.Administrateur
                Return True

            Case RoleWU.Controleur
                If cible.Role = RoleWU.Administrateur Then Return False
                Return Not String.Equals(cible.Identifiant, Identifiant,
                                         StringComparison.OrdinalIgnoreCase)

            Case Else
                Return False
        End Select
    End Function

    ''' <summary>
    ''' Peut-on créer un compte portant ce rôle ?
    '''
    ''' L'ADMINISTRATEUR LES PROPOSE TOUS ; LE CONTRÔLEUR, DEUX. Celui qui crée un compte en
    ''' choisit le mot de passe initial : un contrôleur autorisé à créer un ADMINISTRATEUR se
    ''' créerait le sien et s'y connecterait. Et un contrôleur créé par un contrôleur n'aurait
    ''' plus d'administrateur au-dessus de lui.
    '''
    ''' LE PARAMÈTRE NE S'APPELLE PAS « role » : VB ne distingue pas la casse, et il masquerait
    ''' la propriété Role de l'utilisateur à l'intérieur même de la fonction qui la lit. Le
    ''' Select Case aurait alors porté sur le rôle DEMANDÉ au lieu du rôle de celui qui demande
    ''' — et l'administrateur n'aurait plus pu créer personne, sans qu'aucun compilateur ne
    ''' s'en plaigne.
    ''' </summary>
    Public Function PeutCreerLeRole(roleDemande As RoleWU) As Boolean

        If roleDemande = RoleWU.Inconnu Then Return False

        Select Case Role
            Case RoleWU.Administrateur
                Return True

            Case RoleWU.Controleur
                Return roleDemande = RoleWU.Compense OrElse roleDemande = RoleWU.Commercial

            Case Else
                Return False
        End Select
    End Function

    ''' <summary>Consulter les rapports d'activité : tout rôle connu y a accès.</summary>
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
            Case ROLE_CONTROLEUR : Return RoleWU.Controleur
            Case Else : Return RoleWU.Inconnu
        End Select
    End Function

    ''' <summary>Libellé enregistré correspondant au rôle.</summary>
    Public Shared Function LibelleDepuisRole(role As RoleWU) As String

        Select Case role
            Case RoleWU.Compense : Return ROLE_COMPENSE
            Case RoleWU.Commercial : Return ROLE_COMMERCIAL
            Case RoleWU.Administrateur : Return ROLE_ADMIN
            Case RoleWU.Controleur : Return ROLE_CONTROLEUR
            Case Else : Return String.Empty
        End Select
    End Function

    ''' <summary>Libellé lisible d'un rôle, pour l'affichage et les listes déroulantes.</summary>
    Public Shared Function LibelleLisibleDepuisRole(role As RoleWU) As String

        Select Case role
            Case RoleWU.Compense : Return "Agent de la compense"
            Case RoleWU.Commercial : Return "Commercial"
            Case RoleWU.Administrateur : Return "Administrateur"
            Case RoleWU.Controleur : Return "Contrôleur"
            Case Else : Return "Rôle inconnu"
        End Select
    End Function

    ''' <summary>Libellé lisible du rôle de cet utilisateur.</summary>
    Public ReadOnly Property LibelleRole As String
        Get
            Return LibelleLisibleDepuisRole(Role)
        End Get
    End Property

    ''' <summary>Libellés lisibles des quatre rôles, dans l'ordre où ils sont proposés.</summary>
    Public Shared ReadOnly Property RolesProposes As String()
        Get
            Return New String() {"Agent de la compense", "Commercial", "Administrateur", "Contrôleur"}
        End Get
    End Property

    ''' <summary>
    ''' Libellés des rôles que CET utilisateur a le droit de créer, dans le même ordre.
    '''
    ''' LA LISTE DÉROULANTE NE PROPOSE QUE CE QUI SERA ACCEPTÉ. Afficher à un contrôleur un
    ''' rôle que le dépôt refusera ensuite lui ferait saisir un compte entier pour s'entendre
    ''' dire non à l'enregistrement — et laisserait croire que la règle est négociable.
    ''' </summary>
    Public Shared Function RolesProposesPar(createur As UtilisateurWU) As String()

        Dim liste As New List(Of String)()

        For Each libelle As String In RolesProposes
            Dim role As RoleWU = RoleDepuisLibelleLisible(libelle)
            If createur IsNot Nothing AndAlso createur.PeutCreerLeRole(role) Then
                liste.Add(libelle)
            End If
        Next

        Return liste.ToArray()
    End Function

    ''' <summary>Fonction correspondant au libellé enregistré.</summary>
    Public Shared Function FonctionDepuisLibelle(libelle As String) As FonctionWU

        Select Case If(libelle, String.Empty).Trim().ToUpperInvariant()
            Case FONCTION_INPUTER : Return FonctionWU.Inputer
            Case FONCTION_AUTHORIZER : Return FonctionWU.Authorizer
            Case Else : Return FonctionWU.Aucune
        End Select
    End Function

    ''' <summary>Libellé enregistré correspondant à la fonction. Vide pour « aucune ».</summary>
    Public Shared Function LibelleDepuisFonction(fonction As FonctionWU) As String

        Select Case fonction
            Case FonctionWU.Inputer : Return FONCTION_INPUTER
            Case FonctionWU.Authorizer : Return FONCTION_AUTHORIZER
            Case Else : Return String.Empty
        End Select
    End Function

    ''' <summary>Libellé lisible d'une fonction.</summary>
    Public Shared Function LibelleLisibleDepuisFonction(fonction As FonctionWU) As String

        Select Case fonction
            Case FonctionWU.Inputer : Return "Inputer (saisit)"
            Case FonctionWU.Authorizer : Return "Authorizer (autorise)"
            Case Else : Return "Aucune"
        End Select
    End Function

    ''' <summary>Libellé lisible de la fonction de cet utilisateur.</summary>
    Public ReadOnly Property LibelleFonction As String
        Get
            Return LibelleLisibleDepuisFonction(Fonction)
        End Get
    End Property

    ''' <summary>Libellés lisibles des fonctions, dans l'ordre où elles sont proposées.</summary>
    Public Shared ReadOnly Property FonctionsProposees As String()
        Get
            Return New String() {"Aucune", "Inputer (saisit)", "Authorizer (autorise)"}
        End Get
    End Property

    ''' <summary>Fonction correspondant à un libellé lisible (celui d'une liste déroulante).</summary>
    Public Shared Function FonctionDepuisLibelleLisible(libelle As String) As FonctionWU

        Select Case If(libelle, String.Empty).Trim()
            Case "Inputer (saisit)" : Return FonctionWU.Inputer
            Case "Authorizer (autorise)" : Return FonctionWU.Authorizer
            Case Else : Return FonctionWU.Aucune
        End Select
    End Function

    ''' <summary>Rôle correspondant à un libellé lisible (celui d'une liste déroulante).</summary>
    Public Shared Function RoleDepuisLibelleLisible(libelle As String) As RoleWU

        Select Case If(libelle, String.Empty).Trim()
            Case "Agent de la compense" : Return RoleWU.Compense
            Case "Commercial" : Return RoleWU.Commercial
            Case "Administrateur" : Return RoleWU.Administrateur
            Case "Contrôleur" : Return RoleWU.Controleur
            Case Else : Return RoleWU.Inconnu
        End Select
    End Function

#End Region

End Class
