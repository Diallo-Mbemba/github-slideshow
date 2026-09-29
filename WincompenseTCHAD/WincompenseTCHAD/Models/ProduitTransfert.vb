Option Strict On
Option Explicit On

Imports System.Drawing
Imports System.Text.RegularExpressions

''' <summary>
''' Un produit de transfert d'argent traité par Wincompense.
'''
''' POURQUOI CETTE CLASSE EXISTE
'''
''' L'application a longtemps été « l'application Western Union » : son titre le disait, ses
''' tables portent encore son nom. Elle est devenue l'application de compensation des
''' transferts d'argent de la banque, dont Western Union n'est que le premier produit.
'''
''' LA RÈGLE QUI TIENT TOUT LE RESTE
'''
''' Aucun écran ne demande jamais « suis-je Western Union ? ». Tout ce qui dépend du produit —
''' son nom, son état, les écrans qu'il propose, et demain le jeu de tables qu'il lit — est
''' porté par CET objet et par lui seul.
'''
''' UN PRODUIT A DEUX MOITIÉS, ET UNE SEULE SE SAISIT
'''
''' La première version écrivait la liste en dur, au motif qu'une table donnerait l'illusion
''' qu'on ajoute un produit en saisissant une ligne. L'argument ne valait que pour une moitié :
'''
'''   CE QUI EST UNE DONNÉE : l'existence du produit, son nom, sa description, sa couleur, son
'''   ordre d'affichage, son état de service. Cela se saisit, et l'administrateur le fait.
'''
'''   CE QUI EST DU CODE : le format de ses rapports, ses taux, sa cascade de commissions, sa
'''   pièce comptable, son jeu de tables. Cela se livre.
'''
''' D'où la règle de sécurité, qui est le cœur de cette classe :
'''
'''   LA BASE DIT QUELS PRODUITS EXISTENT. LE CODE DIT LESQUELS SONT TRAITÉS.
'''   L'ADMINISTRATEUR NE PEUT PAS TOUCHER À LA SECONDE.
'''
''' Sans elle, un administrateur déclarant MoneyGram « disponible » ouvrirait l'écran de
''' traitement de la compense — qui lirait les rapports Western Union, les taux Western Union,
''' et produirait une pièce sur les comptes Western Union, SOUS UN NOM MONEYGRAM. Personne ne
''' s'en apercevrait avant la comptabilisation.
''' </summary>
Public NotInheritable Class ProduitTransfert

#Region "Identité de l'application"

    ''' <summary>
    ''' Le nom du logiciel et la filiale qu'il traite.
    '''
    ''' « Wincompense » est le nom du logiciel ; « Compensation Western Union » ne l'a jamais
    ''' été — c'était le nom de son seul produit.
    ''' </summary>
    Public Const ENSEIGNE As String = "Wincompense TCHAD"

#End Region

#Region "Ce que le code sait traiter"

    Public Const CODE_WESTERN_UNION As String = "WU"
    Public Const CODE_RIA As String = "RIA"

    ''' <summary>
    ''' Les produits dont le traitement est ÉCRIT dans cette version de l'application.
    '''
    ''' CETTE LISTE N'EST PAS MODIFIABLE DEPUIS L'INTERFACE, et c'est tout son intérêt. Elle ne
    ''' grandit qu'avec une livraison qui apporte, pour le produit ajouté, la lecture de ses
    ''' rapports, ses taux, sa pièce comptable et son jeu de tables.
    '''
    ''' Une propriété et non un champ : les champs Shared s'initialisent dans l'ordre de leur
    ''' déclaration, et ce piège a déjà coûté une livraison ici. Voir verif_initialisation.py.
    ''' </summary>
    Private Shared ReadOnly Property CodesTraites As String()
        Get
            Return New String() {CODE_WESTERN_UNION}
        End Get
    End Property

    ''' <summary>Vrai si le traitement de ce produit est écrit dans cette version.</summary>
    Public Shared Function EstTraiteParLApplication(code As String) As Boolean

        If String.IsNullOrWhiteSpace(code) Then Return False

        For Each connu As String In CodesTraites
            If String.Equals(connu, code.Trim(), StringComparison.OrdinalIgnoreCase) Then Return True
        Next

        Return False
    End Function

#End Region

#Region "La liste des produits"

    Private Shared _tous As IList(Of ProduitTransfert) = ListeDeSecours()

    ''' <summary>Tous les produits connus, en service ou non, dans l'ordre d'affichage.</summary>
    Public Shared ReadOnly Property Tous As IList(Of ProduitTransfert)
        Get
            Return _tous
        End Get
    End Property

    ''' <summary>
    ''' Les seuls produits que l'utilisateur voit à la connexion.
    '''
    ''' Le nom ne peut pas être « EnService » tout court : la propriété d'instance porte déjà
    ''' ce nom, et VB ne distingue pas la casse -- deux membres homonymes dans une même classe
    ''' sont refusés, Shared ou non.
    ''' </summary>
    Public Shared Function ProduitsEnService() As IList(Of ProduitTransfert)

        Dim retenus As New List(Of ProduitTransfert)

        For Each produit As ProduitTransfert In _tous
            If produit.EnService Then retenus.Add(produit)
        Next

        Return retenus
    End Function

    ''' <summary>
    ''' Remplace la liste par celle que l'administrateur a saisie, lue en base.
    '''
    ''' DEUX GARDE-FOUS, QUI NE SONT PAS DE LA POLITESSE
    '''
    ''' 1. Un produit que le CODE sait traiter est ajouté d'office s'il manque. Une ligne
    '''    supprimée à la main, un script non passé, une base restaurée d'avant : Western Union
    '''    doit rester ouvrable, sans quoi la banque ne peut plus compenser.
    '''
    ''' 2. Si plus aucun produit n'est en service, ceux que le code traite y sont remis. Une
    '''    liste vide, c'est une application où personne ne peut plus travailler, et il n'y
    '''    aurait plus d'écran pour revenir en arrière.
    ''' </summary>
    Public Shared Sub Charger(lus As IEnumerable(Of ProduitTransfert))

        Dim liste As New List(Of ProduitTransfert)
        If lus IsNot Nothing Then liste.AddRange(lus)

        ' --- garde-fou 1 : les produits traités par le code ne peuvent pas disparaître
        For Each secours As ProduitTransfert In ListeDeSecours()

            If Not EstTraiteParLApplication(secours.Code) Then Continue For
            If liste.Any(Function(p) p.PorteLeCode(secours.Code)) Then Continue For
            liste.Add(secours)
        Next

        ' --- garde-fou 2 : au moins un produit en service
        If Not liste.Any(Function(p) p.EnService) Then
            For Each produit As ProduitTransfert In liste
                If produit.EstTraite Then produit.EnService = True
            Next
        End If

        ' Une fonction nommée et non une lambda : Sort accepte un IComparer comme une
        ' Comparison, et une lambda posée là laisserait le choix de la surcharge au
        ' compilateur. AddressOf ne laisse aucun doute.
        liste.Sort(AddressOf ParRangDAffichage)

        _tous = liste
    End Sub

    ''' <summary>
    ''' L'ordre d'affichage : le rang saisi, puis le nom à rang égal. Deux produits posés au
    ''' même rang doivent sortir dans un ordre stable, sans quoi la fenêtre de choix les
    ''' permuterait d'un lancement à l'autre.
    ''' </summary>
    Private Shared Function ParRangDAffichage(gauche As ProduitTransfert,
                                              droite As ProduitTransfert) As Integer

        Dim ecart As Integer = gauche.Ordre.CompareTo(droite.Ordre)
        If ecart <> 0 Then Return ecart

        Return String.Compare(gauche.Nom, droite.Nom, StringComparison.CurrentCulture)
    End Function

    ''' <summary>
    ''' La liste écrite en dur, employée quand la table n'existe pas encore — base sur laquelle
    ''' le script des produits n'est pas passé — ou qu'elle est illisible. L'application doit
    ''' démarrer et compenser sans elle : c'est la règle de la maison, déjà appliquée aux
    ''' comptes systèmes et au barème des taxes.
    ''' </summary>
    Public Shared Function ListeDeSecours() As IList(Of ProduitTransfert)

        Dim occidental As New ProduitTransfert(CODE_WESTERN_UNION)
        occidental.Nom = "Western Union"
        occidental.Description = "La compensation Western Union : rapport d'activité, commissions " &
                                 "sur envois et sur réceptions, pièce comptable et fichier core banking."
        occidental.Couleur = Color.FromArgb(255, 182, 0)
        occidental.Ordre = 10
        occidental.EnService = True

        Dim ria As New ProduitTransfert(CODE_RIA)
        ria.Nom = "Ria"
        ria.Description = "Produit en projet à la banque. L'environnement est en place ; " &
                          "le traitement de sa compensation reste à écrire."
        ria.Couleur = Color.FromArgb(238, 118, 35)
        ria.Ordre = 20
        ria.EnService = True

        Return New List(Of ProduitTransfert) From {occidental, ria}
    End Function

    ''' <summary>
    ''' Vrai si ce code est celui du produit en cours de traitement.
    '''
    ''' La question se pose ICI et nulle part ailleurs : la règle de la maison veut qu'aucun
    ''' écran n'examine le code d'un produit, et verif_produit.py la fait respecter. Un écran
    ''' qui a besoin de savoir « est-ce celui qu'on traite ? » le demande, il ne le déduit pas.
    ''' </summary>
    Public Shared Function EstLeProduitActif(code As String) As Boolean

        If Actif Is Nothing OrElse String.IsNullOrWhiteSpace(code) Then Return False
        Return Actif.PorteLeCode(code)
    End Function

    ''' <summary>Le produit portant ce code, ou Nothing. La casse est ignorée.</summary>
    Public Shared Function ParCode(code As String) As ProduitTransfert

        If String.IsNullOrWhiteSpace(code) Then Return Nothing

        For Each produit As ProduitTransfert In _tous
            If produit.PorteLeCode(code) Then Return produit
        Next

        Return Nothing
    End Function

#End Region

#Region "Le produit en cours"

    ''' <summary>
    ''' Le produit choisi à l'ouverture de l'espace de travail, ou Nothing.
    '''
    ''' Tant qu'aucun produit n'est choisi, ActifEstDisponible vaut Faux et TOUS les écrans
    ''' métier se refusent — même règle que SessionWU, qui refuse tous les droits tant que
    ''' personne n'est connecté.
    ''' </summary>
    Public Shared Property Actif As ProduitTransfert

    ''' <summary>
    ''' Vrai si un produit est choisi ET que son traitement est écrit ET qu'il est en service.
    ''' C'est la question que pose l'espace de travail avant d'ouvrir un écran métier.
    ''' </summary>
    Public Shared ReadOnly Property ActifEstDisponible As Boolean
        Get
            Return Actif IsNot Nothing AndAlso Actif.Disponible
        End Get
    End Property

    ''' <summary>Nom du produit en cours, ou une chaîne vide. Sert aux messages.</summary>
    Public Shared ReadOnly Property NomDuProduitActif As String
        Get
            Return If(Actif Is Nothing, String.Empty, Actif.Nom)
        End Get
    End Property

    ''' <summary>
    ''' Titre de l'espace de travail : l'enseigne, puis le produit. Sans produit choisi, il
    ''' s'en tient à l'enseigne plutôt que d'afficher un tiret suivi de rien.
    ''' </summary>
    Public Shared ReadOnly Property TitreDeLEspace As String
        Get
            If Actif Is Nothing Then Return ENSEIGNE
            Return $"{ENSEIGNE} — {Actif.Nom}"
        End Get
    End Property

#End Region

#Region "Le code d'un produit"

    ''' <summary>Longueurs admises pour un code de produit.</summary>
    Public Const CODE_LONGUEUR_MINIMALE As Integer = 2
    Public Const CODE_LONGUEUR_MAXIMALE As Integer = 10

    ''' <summary>
    ''' La forme admise, hors longueur : une lettre, puis des lettres et des chiffres. La
    ''' longueur est vérifiée à part, pour que le message dise laquelle des deux règles est
    ''' en cause.
    ''' </summary>
    Private Shared ReadOnly Property FormeDuCode As Regex
        Get
            Return New Regex("^[A-Z][A-Z0-9]*$")
        End Get
    End Property

    ''' <summary>
    ''' Vérifie qu'un code peut servir de code de produit, et dit pourquoi s'il ne peut pas.
    '''
    ''' LES CONTRAINTES NE SONT PAS COSMÉTIQUES. Le code nomme le fichier du logo
    ''' (Logos\MGRAM.png) et nommera demain le schéma ou le suffixe des tables du produit. Un
    ''' espace, un accent ou une barre oblique y deviendraient un chemin invalide ou un nom
    ''' d'objet SQL à quoter.
    ''' </summary>
    Public Shared Function CodeValide(code As String, ByRef erreur As String) As Boolean

        erreur = String.Empty

        If String.IsNullOrWhiteSpace(code) Then
            erreur = "Le code du produit est obligatoire."
            Return False
        End If

        Dim propre As String = code.Trim().ToUpperInvariant()

        If propre.Length < CODE_LONGUEUR_MINIMALE OrElse propre.Length > CODE_LONGUEUR_MAXIMALE Then
            erreur = $"Le code doit compter entre {CODE_LONGUEUR_MINIMALE} et " &
                     $"{CODE_LONGUEUR_MAXIMALE} caractères."
            Return False
        End If

        If Not FormeDuCode.IsMatch(propre) Then
            erreur = "Le code ne peut contenir que des lettres non accentuées et des chiffres, " &
                     "et doit commencer par une lettre. Il nomme le fichier du logo et, demain, " &
                     "les tables du produit."
            Return False
        End If

        Return True
    End Function

#End Region

#Region "Un produit"

    ''' <summary>
    ''' Crée un produit à partir de son code, seul trait qui ne changera plus. Le code nomme le
    ''' fichier du logo et, demain, les tables du produit : le modifier après coup orphelinerait
    ''' les unes et les autres.
    ''' </summary>
    Public Sub New(code As String)
        _code = If(code Is Nothing, String.Empty, code.Trim().ToUpperInvariant())
        _nom = _code
        _description = String.Empty
        _couleur = COULEUR_PAR_DEFAUT
        _ordre = 100
        _enService = True
    End Sub

    ''' <summary>Gris neutre d'un produit dont l'administrateur n'a pas choisi la couleur.</summary>
    Private Shared ReadOnly Property COULEUR_PAR_DEFAUT As Color
        Get
            Return Color.FromArgb(108, 117, 125)
        End Get
    End Property

    Private ReadOnly _code As String
    Private _nom As String
    Private _description As String
    Private _couleur As Color
    Private _ordre As Integer
    Private _enService As Boolean

    ''' <summary>Code court du produit — « WU », « RIA ». Il nomme son logo et ses tables.</summary>
    Public ReadOnly Property Code As String
        Get
            Return _code
        End Get
    End Property

    ''' <summary>Vrai si ce produit porte ce code. La casse et les espaces sont ignorés.</summary>
    Public Function PorteLeCode(code As String) As Boolean

        If String.IsNullOrWhiteSpace(code) Then Return False
        Return String.Equals(_code, code.Trim(), StringComparison.OrdinalIgnoreCase)
    End Function

    ''' <summary>Nom affiché du produit. Librement corrigeable, contrairement au code.</summary>
    Public Property Nom As String
        Get
            Return _nom
        End Get
        Set(valeur As String)
            _nom = If(String.IsNullOrWhiteSpace(valeur), _code, valeur.Trim())
        End Set
    End Property

    ''' <summary>Ce que le produit recouvre, en une phrase, pour la fenêtre de choix.</summary>
    Public Property Description As String
        Get
            Return _description
        End Get
        Set(valeur As String)
            _description = If(valeur Is Nothing, String.Empty, valeur.Trim())
        End Set
    End Property

    ''' <summary>Fond de l'emblème dessiné à défaut du logo officiel.</summary>
    Public Property Couleur As Color
        Get
            Return _couleur
        End Get
        Set(valeur As Color)
            ' Une couleur vide donnerait un pinceau entièrement transparent, et un emblème sans
            ' fond. Le piège a déjà frappé : il ne passera pas deux fois.
            _couleur = If(valeur.A = 0, COULEUR_PAR_DEFAUT, Color.FromArgb(255, valeur))
        End Set
    End Property

    ''' <summary>Rang d'affichage dans la fenêtre de choix, croissant.</summary>
    Public Property Ordre As Integer
        Get
            Return _ordre
        End Get
        Set(valeur As Integer)
            _ordre = valeur
        End Set
    End Property

    ''' <summary>
    ''' Faux si l'administrateur a mis le produit hors service : il n'apparaît alors plus dans
    ''' la fenêtre de choix. Cela ne supprime rien — l'historique du produit reste en place.
    ''' </summary>
    Public Property EnService As Boolean
        Get
            Return _enService
        End Get
        Set(valeur As Boolean)
            _enService = valeur
        End Set
    End Property

    ''' <summary>Traçabilité, lue et écrite par le dépôt.</summary>
    Public Property CreePar As String
    Public Property DateCreation As Date?
    Public Property ModifiePar As String
    Public Property DateModification As Date?

    ''' <summary>
    ''' Vrai si le traitement de ce produit est écrit dans cette version de l'application.
    ''' AUCUN ÉCRAN NE PEUT CHANGER CELA : c'est une propriété du code, pas de la donnée.
    ''' </summary>
    Public ReadOnly Property EstTraite As Boolean
        Get
            Return EstTraiteParLApplication(_code)
        End Get
    End Property

    ''' <summary>
    ''' Vrai si l'espace de travail de ce produit ouvre de vrais écrans. Les deux conditions
    ''' sont nécessaires, et l'une d'elles n'est pas saisissable.
    ''' </summary>
    Public ReadOnly Property Disponible As Boolean
        Get
            Return EstTraite AndAlso _enService
        End Get
    End Property

    ''' <summary>État du produit, tel que les écrans l'écrivent.</summary>
    Public ReadOnly Property Etat As String
        Get
            If Not _enService Then Return "Hors service"
            If Not EstTraite Then Return "En attente"
            Return "Disponible"
        End Get
    End Property

    ''' <summary>
    ''' Couleur des initiales sur l'emblème : noir ou blanc, CALCULÉE sur la luminance du fond.
    '''
    ''' Elle n'est pas saisie, et c'est voulu. L'administrateur qui choisit librement les deux
    ''' finit un jour par poser du jaune sur du blanc, et l'emblème devient illisible sans que
    ''' rien ne l'en avertisse. La formule est celle de la luminance perçue : l'œil est bien
    ''' plus sensible au vert qu'au bleu.
    ''' </summary>
    Public ReadOnly Property CouleurDuTexte As Color
        Get
            Dim luminance As Double = 0.299R * _couleur.R + 0.587R * _couleur.G + 0.114R * _couleur.B
            Return If(luminance > 150.0R, Color.Black, Color.White)
        End Get
    End Property

    ''' <summary>
    ''' Les initiales portées par l'emblème. C'est le code du produit : il est court par
    ''' construction, et c'est déjà ainsi que la banque nomme ses produits.
    ''' </summary>
    Public ReadOnly Property Initiales As String
        Get
            Return _code
        End Get
    End Property

    Public Overrides Function ToString() As String
        Return _nom
    End Function

#End Region

End Class
