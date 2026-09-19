Option Strict On
Option Explicit On

Imports System.Configuration
Imports System.IO
Imports System.Text

''' <summary>
''' Trouve la chaîne de connexion à SQL Server, et sait la changer.
'''
''' La banque change souvent de serveur. Tant que la chaîne vivait dans le fichier de
''' configuration posé à côté de l'exécutable, en changer imposait une tournée dans les bureaux :
''' ce fichier est dans Program Files, donc protégé, il est propre à chaque poste, et une
''' réinstallation l'écrase.
'''
''' La chaîne est donc cherchée ailleurs, dans un ordre où le premier trouvé l'emporte :
'''
'''   1. la variable d'environnement WINCOMPENSE_CONNEXION — dépannage, poste de test ;
'''   2. le FICHIER PARTAGÉ, dont le chemin est inscrit sur le poste à l'installation :
'''      c'est la source de vérité, et le seul fichier à changer pour toute la banque ;
'''   3. le fichier local %PROGRAMDATA%\Wincompense\wincompense.config, rafraîchi à chaque
'''      lecture réussie du partage ;
'''   4. App.config, pour un poste de développement qui n'a rien installé ;
'''   5. la valeur compilée, dernier recours.
'''
''' Le rang 3 n'est pas un doublon : c'est ce qui permet de travailler le matin où le partage
''' est injoignable. Sans lui, une coupure réseau arrêterait la compense de toute la banque.
'''
''' Format des deux fichiers : une ligne CLE=VALEUR, le dièse ouvre un commentaire. Ni XML ni
''' registre — l'informatique de la banque doit pouvoir changer de serveur dans le Bloc-notes,
''' sans outil et sans risque de casser une balise.
''' </summary>
Public NotInheritable Class ConfigurationWU

    Private Sub New()
    End Sub

#Region "Emplacements et clés"

    ''' <summary>Nom du dossier de l'application sous %PROGRAMDATA%.</summary>
    Private Const DOSSIER As String = "Wincompense"

    ''' <summary>Fichier local : chemin du partage, et copie de secours de la connexion.</summary>
    Private Const FICHIER_LOCAL As String = "wincompense.config"

    Private Const VARIABLE_ENVIRONNEMENT As String = "WINCOMPENSE_CONNEXION"

    ''' <summary>Clé du fichier local portant le chemin du fichier partagé.</summary>
    Public Const CLE_PARTAGE As String = "PARTAGE"

    Public Const CLE_SERVEUR As String = "SERVEUR"
    Public Const CLE_BASE As String = "BASE"
    Public Const CLE_DELAI As String = "DELAI"

    ''' <summary>Chaîne complète, pour les cas exceptionnels : elle prime sur SERVEUR/BASE/DELAI.</summary>
    Public Const CLE_CHAINE As String = "CHAINE"

    ''' <summary>
    ''' Mot de passe du compte SQL Server, chiffré pour ce poste par <see cref="SecretWU"/>.
    '''
    ''' Il ne vit QUE dans le fichier local, jamais sur le partage : le partage est lisible par
    ''' tous les utilisateurs de l'application, et c'est précisément ce qui lui permet de valoir
    ''' pour tout le monde. Un secret n'y a donc pas sa place, chiffré ou non — chiffré pour un
    ''' poste, il serait de toute façon illisible sur les autres.
    ''' </summary>
    Public Const CLE_MOTDEPASSE As String = "MOTDEPASSE"

    ''' <summary>
    ''' Mot de passe encore en clair, écrit par le programme d'installation.
    '''
    ''' L'installateur ne sait pas chiffrer pour Windows ; l'application, si. Cette clé est donc
    ''' un passage : à la première lecture, elle est chiffrée sous MOTDEPASSE puis EFFACÉE. La
    ''' fenêtre pendant laquelle le mot de passe est en clair se réduit à l'intervalle entre
    ''' l'installation et le premier démarrage.
    ''' </summary>
    Public Const CLE_MOTDEPASSE_CLAIR As String = "MOTDEPASSE_CLAIR"

    Public Const BASE_PAR_DEFAUT As String = "GWC_WINCOMPENSE_ETD"
    Public Const DELAI_PAR_DEFAUT As Integer = 10

    ''' <summary>Nom de la section App.config conservée pour les postes de développement.</summary>
    Private Const SECTION_APP_CONFIG As String = BASE_PAR_DEFAUT

    ''' <summary>Dossier local de configuration : %PROGRAMDATA%\Wincompense.</summary>
    Public Shared ReadOnly Property DossierLocal As String
        Get
            Return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), DOSSIER)
        End Get
    End Property

    ''' <summary>Fichier local de configuration, qu'il existe ou non.</summary>
    Public Shared ReadOnly Property CheminLocal As String
        Get
            Return Path.Combine(DossierLocal, FICHIER_LOCAL)
        End Get
    End Property

    ''' <summary>
    ''' Dossier du fichier partagé, ou chaîne vide si ce poste n'en a pas.
    '''
    ''' C'est là que se rangent les fichiers communs à toute la banque : la connexion, et la
    ''' version publiée. Les déduire du chemin déjà connu évite un second réglage à tenir à
    ''' jour — et un second réglage qu'on oublierait de changer le jour où le partage bouge.
    ''' </summary>
    Public Shared ReadOnly Property DossierDuPartage As String
        Get
            Dim fichier As String = CheminDuPartage
            If fichier.Length = 0 Then Return String.Empty

            Try
                Return If(Path.GetDirectoryName(fichier), String.Empty)
            Catch
                ' Chemin mal formé : mieux vaut ne rien proposer que de composer un chemin faux.
                Return String.Empty
            End Try
        End Get
    End Property

    ''' <summary>
    ''' Chaîne complète inscrite dans la configuration en service (clé CHAINE), ou chaîne vide.
    '''
    ''' La banque fournit parfois une chaîne toute faite, avec des mots-clés que SERVEUR, BASE
    ''' et DELAI ne savent pas exprimer : chiffrement imposé, partenaire de secours, nom
    ''' d'application. L'écran d'administration a besoin de savoir qu'une telle chaîne est en
    ''' place, faute de quoi il l'écraserait en enregistrant trois valeurs décomposées.
    ''' </summary>
    Public Shared ReadOnly Property ChaineComplete As String
        Get
            Dim local As Dictionary(Of String, String) = LireFichier(CheminLocal)
            Dim partage As String = LireCle(local, CLE_PARTAGE)

            ' Même ordre que la résolution : le partage commande, le local suit.
            If partage.Length > 0 Then
                Dim valeurs As Dictionary(Of String, String) = LireFichier(partage)
                If valeurs.Count > 0 Then Return LireCle(valeurs, CLE_CHAINE)
            End If

            Return LireCle(local, CLE_CHAINE)
        End Get
    End Property

    ''' <summary>
    ''' Chemin du fichier partagé, tel qu'il est inscrit sur ce poste. Chaîne vide si le poste
    ''' n'a pas été installé, ou si l'installation n'a pas désigné de partage.
    ''' </summary>
    Public Shared ReadOnly Property CheminDuPartage As String
        Get
            Return LireCle(LireFichier(CheminLocal), CLE_PARTAGE)
        End Get
    End Property

#End Region

#Region "Résolution de la chaîne de connexion"

    ''' <summary>Chaîne retenue, calculée une fois puis conservée le temps de l'exécution.</summary>
    Private Shared _chaine As String = String.Empty

    ''' <summary>
    ''' Chaîne imposée pour cette exécution seulement, sans rien écrire sur le poste.
    '''
    ''' Elle sert au technicien qui installe : il peut désigner la bonne base le temps de créer
    ''' le premier administrateur, sans que le poste cesse pour autant de suivre la chaîne
    ''' publiée avec l'application. Un réglage écrit sur le poste l'emporterait en effet sur
    ''' App.config — définitivement, et sans que rien ne le signale.
    ''' </summary>
    Private Shared _forcageSession As String = String.Empty

    ''' <summary>Provenance de la chaîne retenue, en clair, pour l'écran d'administration.</summary>
    Private Shared _origine As String = String.Empty

    ''' <summary>
    ''' Chaîne de connexion à employer. Ne lève jamais : un poste mal configuré doit atteindre
    ''' l'écran de connexion et y lire un message, non se fermer sur une exception.
    ''' </summary>
    Public Shared Function ChaineDeConnexion() As String

        If _chaine.Length > 0 Then Return _chaine

        Resoudre()
        Return _chaine
    End Function

    ''' <summary>
    ''' D'où vient la chaîne en service : « fichier partagé », « copie locale », etc.
    ''' Affiché à l'administrateur, qui doit pouvoir constater que le partage est bien lu.
    ''' </summary>
    Public Shared Function Origine() As String

        If _chaine.Length = 0 Then Resoudre()
        Return _origine
    End Function

    ''' <summary>
    ''' Oblige à tout relire au prochain appel. À employer après un enregistrement, pour que le
    ''' changement prenne effet sans quitter l'application.
    ''' </summary>
    Public Shared Sub Oublier()
        _chaine = String.Empty
        _origine = String.Empty
    End Sub

    ''' <summary>
    ''' Impose une chaîne pour cette exécution, sans rien écrire sur le poste. Au prochain
    ''' démarrage, la résolution normale reprend — et le poste continue donc de suivre la
    ''' chaîne publiée avec l'application.
    ''' </summary>
    ''' <param name="chaine">Chaîne à employer. Chaîne vide pour annuler le forçage.</param>
    Public Shared Sub ForcerPourCetteSession(chaine As String)

        _forcageSession = If(chaine, String.Empty).Trim()
        Oublier()
    End Sub

    ''' <summary>Vrai si la chaîne en service a été saisie pour cette session seulement.</summary>
    Public Shared ReadOnly Property EstForceePourLaSession As Boolean
        Get
            Return _forcageSession.Length > 0
        End Get
    End Property

    ''' <summary>
    ''' Trouve la chaîne, puis y remet le mot de passe conservé sur ce poste.
    '''
    ''' Les deux étapes sont distinctes à dessein : la première dit OÙ est le serveur, et peut
    ''' venir du partage donc de la banque entière ; la seconde dit COMMENT s'y annoncer, et
    ''' n'appartient qu'à cette machine. Les mêler rendrait impossible de changer de serveur
    ''' pour tout le monde sans diffuser un secret à tout le monde.
    ''' </summary>
    Private Shared Sub Resoudre()

        NettoyerLeFichierLocal()
        ResoudreLaSource()
        _chaine = AppliquerLeMotDePasse(_chaine)
    End Sub

    Private Shared Sub ResoudreLaSource()

        ' 0. Saisie imposée pour cette session : elle prime sur tout, et ne survit pas à
        ' la fermeture de l'application.
        If _forcageSession.Length > 0 Then
            _chaine = _forcageSession
            _origine = "saisie pour cette session, non conservée sur le poste"
            Return
        End If

        ' 1. Variable d'environnement : elle passe avant tout, et n'engage que ce poste.
        Dim environnement As String = Nothing
        Try
            environnement = Environment.GetEnvironmentVariable(VARIABLE_ENVIRONNEMENT)
        Catch
            ' Un accès refusé aux variables d'environnement ne doit pas arrêter la résolution.
        End Try

        If Not String.IsNullOrWhiteSpace(environnement) Then
            _chaine = environnement.Trim()
            _origine = $"variable d'environnement {VARIABLE_ENVIRONNEMENT}"
            Return
        End If

        Dim local As Dictionary(Of String, String) = LireFichier(CheminLocal)
        Dim partage As String = LireCle(local, CLE_PARTAGE)

        ' Distingue un partage qui ne répond pas d'un partage qui répond mal : les deux
        ' renvoient à la copie locale, mais ne se corrigent pas au même endroit.
        Dim partageLu As Boolean = False

        ' 2. Fichier partagé : la source de vérité.
        If partage.Length > 0 Then

            Dim valeurs As Dictionary(Of String, String) = LireFichier(partage)
            partageLu = valeurs.Count > 0

            If partageLu Then
                Dim chaine As String = Construire(valeurs)

                If chaine.Length > 0 Then
                    _chaine = chaine
                    _origine = $"fichier partagé {partage}"

                    ' La copie locale est rafraîchie maintenant, pendant que le partage répond :
                    ' c'est elle qui fera tourner le poste le jour où il ne répondra plus.
                    RafraichirLaCopieLocale(local, valeurs)
                    Return
                End If
            End If
        End If

        ' 3. Copie locale.
        Dim chaineLocale As String = Construire(local)

        If chaineLocale.Length > 0 Then
            _chaine = chaineLocale
            If partage.Length = 0 Then
                _origine = "fichier local de ce poste"
            ElseIf partageLu Then
                _origine = $"copie locale — {partage} est lisible mais n'indique aucun serveur"
            Else
                _origine = $"copie locale — {partage} injoignable"
            End If

            Return
        End If

        ' 4. App.config, pour un poste de développement.
        Try
            Dim section As ConnectionStringSettings = ConfigurationManager.ConnectionStrings(SECTION_APP_CONFIG)

            If section IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(section.ConnectionString) Then
                _chaine = section.ConnectionString
                _origine = "App.config de l'application"
                Return
            End If
        Catch
            ' Un fichier de configuration illisible retombe simplement sur la valeur par défaut.
        End Try

        ' 5. Dernier recours.
        _chaine = ChaineDepuis(".\SQLEXPRESS", BASE_PAR_DEFAUT, DELAI_PAR_DEFAUT)
        _origine = "valeur par défaut (aucune configuration trouvée)"
    End Sub

    ''' <summary>
    ''' Remet dans la chaîne le mot de passe gardé sur ce poste, quand elle nomme un compte
    ''' SQL Server sans le porter.
    '''
    ''' Rien n'est touché en authentification Windows — aucun compte n'y est nommé — ni quand
    ''' la chaîne porte déjà son mot de passe : ce qui a été saisi explicitement prime toujours
    ''' sur ce que le poste a retenu.
    ''' </summary>
    Private Shared Function AppliquerLeMotDePasse(chaine As String) As String

        If String.IsNullOrWhiteSpace(chaine) Then Return chaine

        Dim constructeur As System.Data.SqlClient.SqlConnectionStringBuilder

        Try
            constructeur = New System.Data.SqlClient.SqlConnectionStringBuilder(chaine)
        Catch ex As ArgumentException
            ' Chaîne inexploitable : l'ouverture de la connexion le dira mieux que nous.
            Return chaine
        End Try

        ' Aucun compte nommé : la chaîne est en authentification Windows, il n'y a pas de mot
        ' de passe à y mettre.
        If constructeur.UserID.Trim().Length = 0 Then Return chaine

        ' Mot de passe déjà présent : on ne le remplace pas.
        If constructeur.Password.Length > 0 Then Return chaine

        Dim motDePasse As String = MotDePasseDuPoste()
        If motDePasse.Length = 0 Then Return chaine

        constructeur.Password = motDePasse
        constructeur.IntegratedSecurity = False

        Return constructeur.ConnectionString
    End Function

    ''' <summary>Mot de passe conservé sur ce poste, déchiffré.</summary>
    Private Shared Function MotDePasseDuPoste() As String
        Return SecretWU.Lire(LireCle(LireFichier(CheminLocal), CLE_MOTDEPASSE))
    End Function

    ''' <summary>
    ''' Met à l'abri tout mot de passe qui traînerait en clair dans le fichier local, puis
    ''' l'en efface. Appelé avant chaque résolution.
    '''
    ''' Deux chemins y déposent du clair, et aucun n'est une faute :
    '''
    '''   — le programme d'installation, qui ne sait pas chiffrer pour Windows et écrit donc
    '''     MOTDEPASSE_CLAIR, à charge pour nous de le reprendre ;
    '''   — une chaîne CHAINE recopiée à la main dans le fichier, mot de passe compris, par
    '''     une informatique qui suit sa propre procédure.
    '''
    ''' Dans les deux cas le mot de passe passe sous MOTDEPASSE, chiffré, et disparaît de sa
    ''' forme lisible. La fenêtre d'exposition se referme au premier démarrage.
    ''' </summary>
    Private Shared Sub NettoyerLeFichierLocal()

        Try
            Dim local As Dictionary(Of String, String) = LireFichier(CheminLocal)
            If local.Count = 0 Then Return

            Dim aEcrire As Boolean = False
            Dim trouve As String = LireCle(local, CLE_MOTDEPASSE_CLAIR)

            If trouve.Length > 0 Then
                local.Remove(CLE_MOTDEPASSE_CLAIR)
                aEcrire = True
            End If

            ' Un mot de passe dans CHAINE prime sur celui déjà rangé : il vient d'être posé là
            ' par quelqu'un, donc il est le plus récent des deux.
            Dim chaine As String = LireCle(local, CLE_CHAINE)

            If chaine.Length > 0 Then
                Dim dansLaChaine As String = ExtraireLeMotDePasse(chaine)

                If dansLaChaine.Length > 0 Then
                    trouve = dansLaChaine
                    local(CLE_CHAINE) = ChaineSansMotDePasse(chaine)
                    aEcrire = True
                End If
            End If

            If Not aEcrire Then Return

            If trouve.Length > 0 Then
                Dim protege As String = SecretWU.Proteger(trouve)
                If protege.Length > 0 Then local(CLE_MOTDEPASSE) = protege
            End If

            EcrireFichier(CheminLocal, local, "Configuration de ce poste")

        Catch ex As Exception
            ' Fichier verrouillé ou droit refusé : la connexion se fera quand même, avec ce que
            ' le fichier contient. Rien ici ne doit empêcher l'application de démarrer.
        End Try
    End Sub

    ''' <summary>
    ''' La même chaîne, complétée du mot de passe gardé sur ce poste s'il lui en manque un.
    '''
    ''' Indispensable à l'écran de test : la chaîne affichée ne porte plus son mot de passe —
    ''' il a été mis à l'abri — et l'essayer telle quelle échouerait sur « Login failed », en
    ''' faisant croire à un réglage cassé alors que l'application, elle, se connecte.
    ''' </summary>
    Public Shared Function ChaineEssayable(chaine As String) As String
        Return AppliquerLeMotDePasse(chaine)
    End Function

    ''' <summary>Vrai si une chaîne de connexion porte un mot de passe.</summary>
    Public Shared Function PorteUnMotDePasse(chaine As String) As Boolean

        If String.IsNullOrWhiteSpace(chaine) Then Return False

        Try
            Return New System.Data.SqlClient.SqlConnectionStringBuilder(chaine).Password.Length > 0
        Catch ex As ArgumentException
            ' Chaîne illisible : on s'en tient à la recherche littérale, qui suffit à avertir.
            Return chaine.IndexOf("Password", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                   chaine.IndexOf("Pwd", StringComparison.OrdinalIgnoreCase) >= 0
        End Try
    End Function

#End Region

#Region "Construction de la chaîne"

    ''' <summary>
    ''' Chaîne de connexion formée à partir d'un serveur et d'une base, en authentification
    ''' Windows intégrée. Aucun mot de passe n'y figure, ce qui est précisément ce qui permet
    ''' de poser la configuration sur un partage lisible par tous.
    '''
    ''' Un compte SQL Server ne passe pas par ici : il s'exprime dans une chaîne complète, et
    ''' son mot de passe est mis à part sur le poste (voir CLE_MOTDEPASSE).
    ''' </summary>
    Public Shared Function ChaineDepuis(serveur As String, base As String, delai As Integer) As String

        Dim constructeur As New System.Data.SqlClient.SqlConnectionStringBuilder() With {
            .DataSource = If(serveur, String.Empty).Trim(),
            .InitialCatalog = If(String.IsNullOrWhiteSpace(base), BASE_PAR_DEFAUT, base.Trim()),
            .IntegratedSecurity = True,
            .ConnectTimeout = If(delai > 0, delai, DELAI_PAR_DEFAUT)
        }

        Return constructeur.ConnectionString
    End Function

    ''' <summary>
    ''' Chaîne décrite par un jeu de clés. CHAINE prime, s'il est renseigné ; sinon la chaîne
    ''' est bâtie sur SERVEUR, BASE et DELAI. Chaîne vide si le serveur n'est pas renseigné :
    ''' sans serveur il n'y a rien à tenter, et mieux vaut passer au rang suivant.
    ''' </summary>
    Private Shared Function Construire(valeurs As Dictionary(Of String, String)) As String

        Dim complete As String = LireCle(valeurs, CLE_CHAINE)
        If complete.Length > 0 Then Return complete

        Dim serveur As String = LireCle(valeurs, CLE_SERVEUR)
        If serveur.Length = 0 Then Return String.Empty

        Dim base As String = LireCle(valeurs, CLE_BASE)

        Dim delai As Integer
        If Not Integer.TryParse(LireCle(valeurs, CLE_DELAI), delai) Then delai = DELAI_PAR_DEFAUT

        Return ChaineDepuis(serveur, base, delai)
    End Function

#End Region

#Region "Lecture et écriture des fichiers"

    ''' <summary>
    ''' Lit un fichier CLE=VALEUR. Retourne un dictionnaire vide si le fichier est absent,
    ''' illisible ou sur un partage injoignable — jamais d'exception : chacun des cinq rangs
    ''' doit pouvoir échouer sans emporter les suivants.
    ''' </summary>
    Private Shared Function LireFichier(chemin As String) As Dictionary(Of String, String)

        Dim valeurs As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

        If String.IsNullOrWhiteSpace(chemin) Then Return valeurs

        Try
            If Not File.Exists(chemin) Then Return valeurs

            For Each ligne As String In File.ReadAllLines(chemin, Encoding.UTF8)

                Dim nette As String = ligne.Trim()
                If nette.Length = 0 OrElse nette.StartsWith("#", StringComparison.Ordinal) Then Continue For

                Dim separateur As Integer = nette.IndexOf("="c)
                If separateur <= 0 Then Continue For

                Dim cle As String = nette.Substring(0, separateur).Trim()
                Dim valeur As String = nette.Substring(separateur + 1).Trim()

                If cle.Length > 0 Then valeurs(cle) = valeur
            Next

        Catch
            ' Partage injoignable, droit refusé, fichier verrouillé : le rang suivant prendra
            ' la main. Signaler ici n'aurait pas de sens — la résolution n'a pas encore échoué.
        End Try

        Return valeurs
    End Function

    ''' <summary>Valeur d'une clé, ou chaîne vide.</summary>
    Private Shared Function LireCle(valeurs As Dictionary(Of String, String), cle As String) As String

        Dim valeur As String = Nothing
        If valeurs Is Nothing OrElse Not valeurs.TryGetValue(cle, valeur) Then Return String.Empty

        Return If(valeur, String.Empty).Trim()
    End Function

    ''' <summary>
    ''' Réécrit la copie locale avec ce que le partage vient de donner, en conservant le chemin
    ''' du partage lui-même. Un échec est ignoré : le poste tourne déjà sur la chaîne lue, et
    ''' refuser de démarrer parce qu'un cache n'a pas pu s'écrire serait absurde.
    ''' </summary>
    Private Shared Sub RafraichirLaCopieLocale(local As Dictionary(Of String, String),
                                               valeursDuPartage As Dictionary(Of String, String))

        Try
            Dim fusion As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

            For Each paire As KeyValuePair(Of String, String) In valeursDuPartage
                fusion(paire.Key) = paire.Value
            Next

            ' Le chemin du partage n'appartient qu'au poste : le partage ne se désigne pas
            ' lui-même, sans quoi on ne saurait plus où chercher le jour où il change.
            fusion(CLE_PARTAGE) = LireCle(local, CLE_PARTAGE)

            ' Le mot de passe non plus n'appartient qu'au poste, et le partage n'en porte
            ' jamais : l'écraser avec ce qui vient du réseau le ferait disparaître à chaque
            ' lecture réussie du partage, et le poste cesserait de se connecter sans que
            ' rien n'ait changé de visible.
            Conserver(local, fusion, CLE_MOTDEPASSE)

            EcrireFichier(CheminLocal, fusion, "Copie locale de secours - regeneree automatiquement")

        Catch
            ' Volontairement silencieux : voir le commentaire ci-dessus.
        End Try
    End Sub

    ''' <summary>
    ''' Écrit un fichier CLE=VALEUR, en créant le dossier au besoin.
    ''' </summary>
    ''' <param name="chemin">Fichier à écrire.</param>
    ''' <param name="valeurs">Clés à inscrire.</param>
    ''' <param name="titre">Ligne de commentaire placée en tête.</param>
    Private Shared Sub EcrireFichier(chemin As String, valeurs As Dictionary(Of String, String), titre As String)

        Dim dossier As String = Path.GetDirectoryName(chemin)
        If Not String.IsNullOrEmpty(dossier) AndAlso Not Directory.Exists(dossier) Then
            Directory.CreateDirectory(dossier)
        End If

        Dim contenu As New StringBuilder()
        contenu.AppendLine("# Wincompense TCHAD - " & titre)
        contenu.AppendLine("# Une ligne CLE=VALEUR. Le caractere # ouvre un commentaire.")
        contenu.AppendLine("# Ecrit le " & Date.Now.ToString("dd/MM/yyyy HH:mm"))
        contenu.AppendLine()

        For Each paire As KeyValuePair(Of String, String) In valeurs
            If String.IsNullOrWhiteSpace(paire.Value) Then Continue For
            contenu.AppendLine($"{paire.Key}={paire.Value}")
        Next

        File.WriteAllText(chemin, contenu.ToString(), Encoding.UTF8)
    End Sub

    ''' <summary>
    ''' Reporte une clé de l'ancien jeu de valeurs vers le nouveau, sauf si le nouveau la
    ''' renseigne déjà. Sert aux clés qui n'appartiennent qu'au poste et qu'une réécriture
    ''' venue d'ailleurs ne doit pas emporter.
    ''' </summary>
    Private Shared Sub Conserver(ancien As Dictionary(Of String, String),
                                 nouveau As Dictionary(Of String, String), cle As String)

        If LireCle(nouveau, cle).Length > 0 Then Return

        Dim valeur As String = LireCle(ancien, cle)
        If valeur.Length > 0 Then nouveau(cle) = valeur
    End Sub

#End Region

#Region "Clés libres du fichier local"

    ''' <summary>
    ''' Valeur d'une clé quelconque du fichier local de ce poste, ou chaîne vide.
    ''' Sert aux réglages qui n'appartiennent qu'à la machine — la dernière version signalée,
    ''' par exemple — et qui n'ont donc rien à faire sur le partage.
    ''' </summary>
    Public Shared Function LireValeurLocale(cle As String) As String
        Return LireCle(LireFichier(CheminLocal), cle)
    End Function

    ''' <summary>
    ''' Inscrit une clé dans le fichier local, en CONSERVANT tout ce qui s'y trouve déjà.
    '''
    ''' Le fichier porte la connexion du poste : le réécrire de zéro pour y poser une seule
    ''' valeur le priverait de son serveur. Un échec d'écriture est ignoré — aucun des réglages
    ''' passant par ici ne vaut qu'on interrompe le travail de l'utilisateur.
    ''' </summary>
    Public Shared Sub EcrireValeurLocale(cle As String, valeur As String)

        Try
            Dim valeurs As Dictionary(Of String, String) = LireFichier(CheminLocal)
            valeurs(cle) = If(valeur, String.Empty)

            EcrireFichier(CheminLocal, valeurs, "Configuration de ce poste")

        Catch
            ' Volontairement silencieux : voir le commentaire ci-dessus.
        End Try
    End Sub

#End Region

#Region "Enregistrement depuis l'écran d'administration"

    ''' <summary>
    ''' Enregistre le serveur et la base, sur ce poste et — si demandé — sur le partage.
    ''' </summary>
    ''' <param name="serveur">Nom ou adresse de l'instance SQL Server.</param>
    ''' <param name="base">Nom de la base.</param>
    ''' <param name="delai">Délai de connexion, en secondes.</param>
    ''' <param name="cheminPartage">Fichier partagé. Chaîne vide pour n'écrire que localement.</param>
    ''' <param name="ecrireSurLePartage">Vrai pour propager le changement à toute la banque.</param>
    ''' <param name="messageErreur">Motif de l'échec, le cas échéant.</param>
    ''' <returns>Vrai si tout ce qui était demandé a été écrit.</returns>
    Public Shared Function Enregistrer(serveur As String, base As String, delai As Integer,
                                       cheminPartage As String, ecrireSurLePartage As Boolean,
                                       ByRef messageErreur As String) As Boolean

        messageErreur = String.Empty

        If String.IsNullOrWhiteSpace(serveur) Then
            messageErreur = "Le serveur n'est pas renseigné : sans lui, l'application n'a rien à joindre."
            Return False
        End If

        Dim valeurs As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase) From {
            {CLE_SERVEUR, serveur.Trim()},
            {CLE_BASE, If(String.IsNullOrWhiteSpace(base), BASE_PAR_DEFAUT, base.Trim())},
            {CLE_DELAI, If(delai > 0, delai, DELAI_PAR_DEFAUT).ToString()}
        }

        Return EcrireLaConfiguration(valeurs, cheminPartage, ecrireSurLePartage, messageErreur)
    End Function

    ''' <summary>
    ''' Enregistre une chaîne de connexion complète, telle que la banque l'a fournie.
    '''
    ''' Elle est écrite sous la clé CHAINE, et SERVEUR / BASE / DELAI ne sont pas écrits du
    ''' tout : les laisser à côté ferait coexister deux descriptions du même serveur, dont une
    ''' seule compte, et la lecture du fichier deviendrait trompeuse.
    ''' </summary>
    ''' <param name="chaine">Chaîne de connexion complète.</param>
    ''' <param name="cheminPartage">Fichier partagé. Chaîne vide pour n'écrire que localement.</param>
    ''' <param name="ecrireSurLePartage">Vrai pour propager le changement à toute la banque.</param>
    ''' <param name="messageErreur">Motif de l'échec, le cas échéant.</param>
    ''' <returns>Vrai si tout ce qui était demandé a été écrit.</returns>
    Public Shared Function EnregistrerChaineComplete(chaine As String, cheminPartage As String,
                                                     ecrireSurLePartage As Boolean,
                                                     ByRef messageErreur As String) As Boolean

        messageErreur = String.Empty

        If String.IsNullOrWhiteSpace(chaine) Then
            messageErreur = "La chaîne de connexion est vide."
            Return False
        End If

        ' Une chaîne sur plusieurs lignes serait coupée à la relecture : le fichier se lit ligne
        ' par ligne. Un collage depuis un courriel en apporte facilement une.
        Dim surUneLigne As String = chaine.Replace(vbCr, " ").Replace(vbLf, " ").Trim()

        ' Contrôle de forme avant d'engager toute la banque : une chaîne mal formée refusée ici
        ' vaut mieux qu'une application qui ne démarre plus nulle part.
        Try
            Dim essai As New System.Data.SqlClient.SqlConnectionStringBuilder(surUneLigne)

            If String.IsNullOrWhiteSpace(essai.DataSource) Then
                messageErreur = "Cette chaîne n'indique aucun serveur (mot-clé Server ou Data Source)."
                Return False
            End If

        Catch ex As ArgumentException
            messageErreur = "Cette chaîne de connexion n'est pas exploitable : " & ex.Message
            Return False
        End Try

        ' Le mot de passe est retiré de la chaîne AVANT toute écriture, et conservé chiffré
        ' pour ce poste. C'est ce qui permet de propager le serveur à toute la banque sans
        ' propager le secret : le fichier partagé ne portera qu'un compte nommé, et chaque
        ' poste tiendra son mot de passe de lui-même.
        Dim aPart As String = ExtraireLeMotDePasse(surUneLigne)

        Dim valeurs As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase) From {
            {CLE_CHAINE, If(aPart.Length > 0, ChaineSansMotDePasse(surUneLigne), surUneLigne)}
        }

        If aPart.Length > 0 Then
            Dim protege As String = SecretWU.Proteger(aPart)

            If protege.Length = 0 Then
                messageErreur = "Windows a refusé de chiffrer le mot de passe sur ce poste." &
                                Environment.NewLine &
                                "Rien n'a été enregistré : l'écrire en clair n'est pas une option."
                Return False
            End If

            valeurs(CLE_MOTDEPASSE) = protege
        End If

        If Not EcrireLaConfiguration(valeurs, cheminPartage, ecrireSurLePartage, messageErreur) Then
            Return False
        End If

        ' Retour à l'authentification Windows : le mot de passe gardé n'a plus d'objet. Le
        ' laisser dormir dans le fichier serait un secret conservé pour rien.
        If aPart.Length = 0 AndAlso Not NommeUnCompte(surUneLigne) Then OublierLeMotDePasse()

        Return True
    End Function

    ''' <summary>Vrai si la chaîne nomme un compte SQL Server (User ID).</summary>
    Private Shared Function NommeUnCompte(chaine As String) As Boolean

        Try
            Return New System.Data.SqlClient.SqlConnectionStringBuilder(chaine).UserID.Trim().Length > 0
        Catch ex As ArgumentException
            Return False
        End Try
    End Function

    ''' <summary>Efface le mot de passe conservé sur ce poste.</summary>
    Public Shared Sub OublierLeMotDePasse()

        Try
            Dim local As Dictionary(Of String, String) = LireFichier(CheminLocal)

            ' Les deux retraits sont faits, puis seulement on décide d'écrire : un OrElse
            ' placé ici sauterait le second dès que le premier a trouvé quelque chose.
            Dim protegeRetire As Boolean = local.Remove(CLE_MOTDEPASSE)
            Dim clairRetire As Boolean = local.Remove(CLE_MOTDEPASSE_CLAIR)

            If Not protegeRetire AndAlso Not clairRetire Then Return

            EcrireFichier(CheminLocal, local, "Configuration de ce poste")
            Oublier()

        Catch ex As Exception
            ' Sans conséquence : un mot de passe inutilisé n'empêche pas de travailler.
        End Try
    End Sub

    ''' <summary>Mot de passe porté par une chaîne, ou chaîne vide.</summary>
    Private Shared Function ExtraireLeMotDePasse(chaine As String) As String

        Try
            Return New System.Data.SqlClient.SqlConnectionStringBuilder(chaine).Password
        Catch ex As ArgumentException
            Return String.Empty
        End Try
    End Function

    ''' <summary>La même chaîne, son mot de passe retiré. Le compte, lui, reste : il n'est
    ''' pas un secret, et c'est lui qui dira au poste quel mot de passe appliquer.</summary>
    Private Shared Function ChaineSansMotDePasse(chaine As String) As String

        Try
            Dim constructeur As New System.Data.SqlClient.SqlConnectionStringBuilder(chaine)

            ' Remove et non Password = "" : affecter une chaîne vide laisserait un
            ' « Password= » sans valeur dans le fichier, qui se lit comme un mot de passe vide
            ' et non comme une absence de mot de passe.
            constructeur.Remove("Password")

            Return constructeur.ConnectionString

        Catch ex As ArgumentException
            Return chaine
        End Try
    End Function

    ''' <summary>
    ''' Écrit la configuration, sur le partage puis sur le poste.
    '''
    ''' L'écriture sur le partage est faite EN PREMIER : c'est elle qui peut échouer, faute de
    ''' droits, et il vaut mieux ne rien avoir changé du tout que d'avoir un poste réglé sur un
    ''' serveur que les autres ignorent.
    ''' </summary>
    Private Shared Function EcrireLaConfiguration(valeurs As Dictionary(Of String, String),
                                                  cheminPartage As String, ecrireSurLePartage As Boolean,
                                                  ByRef messageErreur As String) As Boolean

        If ecrireSurLePartage Then

            If String.IsNullOrWhiteSpace(cheminPartage) Then
                messageErreur = "Aucun fichier partagé n'est indiqué : impossible de propager le changement."
                Return False
            End If

            ' Le mot de passe, même chiffré, ne part pas sur le partage. Chiffré pour un
            ' poste, il serait de toute façon illisible sur les autres ; l'y écrire ne
            ' servirait donc à rien, sinon à laisser croire qu'il est diffusé.
            Dim pourLePartage As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

            For Each paire As KeyValuePair(Of String, String) In valeurs
                If String.Equals(paire.Key, CLE_MOTDEPASSE, StringComparison.OrdinalIgnoreCase) Then Continue For
                If String.Equals(paire.Key, CLE_MOTDEPASSE_CLAIR, StringComparison.OrdinalIgnoreCase) Then Continue For
                pourLePartage(paire.Key) = paire.Value
            Next

            Try
                EcrireFichier(cheminPartage, pourLePartage, "Connexion commune a tous les postes")
            Catch ex As Exception
                messageErreur = $"Écriture impossible sur {cheminPartage} : {ex.Message}" & Environment.NewLine &
                                "Vérifiez que vous avez le droit d'écrire sur ce partage. Rien n'a été modifié."
                Return False
            End Try
        End If

        Try
            Dim local As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

            For Each paire As KeyValuePair(Of String, String) In valeurs
                local(paire.Key) = paire.Value
            Next

            local(CLE_PARTAGE) = If(cheminPartage, String.Empty).Trim()

            ' Même raison qu'au rafraîchissement : changer de serveur ne doit pas effacer le
            ' mot de passe du poste. S'il vient d'être ressaisi, il est déjà dans « valeurs »
            ' et Conserver ne le remplace pas.
            Conserver(LireFichier(CheminLocal), local, CLE_MOTDEPASSE)

            EcrireFichier(CheminLocal, local, "Configuration de ce poste")

        Catch ex As Exception
            messageErreur = $"Écriture impossible dans {CheminLocal} : {ex.Message}"
            Return False
        End Try

        ' La prochaine lecture repart de zéro : le changement vaut sans quitter l'application.
        Oublier()
        Return True
    End Function

#End Region

End Class
