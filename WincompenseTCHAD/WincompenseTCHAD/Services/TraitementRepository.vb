Option Strict On
Option Explicit On

Imports System.Data
Imports System.Data.SqlClient
Imports System.Globalization

''' <summary>
''' En-têtes de traitement des journées comptabilisées (table T_TraitementWU), et leur visa.
'''
''' ÉCRIT AVEC L'HISTORIQUE, DANS LA MÊME TRANSACTION
'''
''' Trois documents décrivent la même journée : l'historique, la pièce et l'en-tête de
''' traitement. Ils ne doivent jamais diverger, donc ils s'écrivent ensemble ou pas du tout.
'''
''' RÉPÉTABLE, ET LE VISA NE SURVIT PAS
'''
''' Recomptabiliser une journée remplace son en-tête — et donc efface son visa. C'est voulu :
''' le visa atteste d'un traitement précis, et refaire la journée en produit un autre.
'''
''' Comme les autres dépôts, aucune exception SQL ne remonte à l'interface.
''' </summary>
Public NotInheritable Class TraitementRepository

    Private Sub New()
    End Sub

    Private Const TABLE As String = "T_TraitementWU"

    ''' <summary>Code d'erreur SQL Server signalant une table absente (« Invalid object name »).</summary>
    Private Const ERREUR_TABLE_ABSENTE As Integer = 208

    Public Const MESSAGE_TABLE_ABSENTE As String =
        "La table T_TraitementWU n'existe pas encore dans la base." & vbCrLf & vbCrLf &
        "Exécutez le script Scripts\00_InstallationComplete.sql : il la crée." & vbCrLf &
        "C'est le SEUL script à exécuter — il contient tous les autres — et il peut être " &
        "rejoué sans risque : il ne crée que ce qui manque." & vbCrLf &
        "Tant qu'elle est absente, le bordereau de fin de journée se reconstitue depuis " &
        "l'historique et la pièce, mais sans le nom des rapports Western Union — et aucune " &
        "journée ne peut être visée."

    Private Const COLONNES As String =
        "DateActivite, DateValeur, NumeroLot, FichierActivite, EmpreinteActivite, " &
        "FichierReglement, EmpreinteReglement, NombrePdv, NombreSousAgents, NombreAgences, " &
        "NombreEcartes, NombreEnvois, NombrePaiements, NombreAnnulations, " &
        "TotalDebit, TotalCredit, EcartArrondi, CompteEcart, " &
        "ComptabilisePar, DateComptabilisation, VisePar, DateVisa, CommentaireVisa"

    ''' <summary>
    ''' Vrai si la table porte les deux colonnes de période. Nothing tant que la question n'a
    ''' pas été posée : la distinction compte, Faux signifierait « vérifié, elles n'y sont pas ».
    ''' </summary>
    Private Shared _periodeDisponible As Boolean?

    ''' <summary>Fait reposer la question. À appeler quand la connexion change de base.</summary>
    Public Shared Sub Oublier()
        _periodeDisponible = Nothing
    End Sub

    ''' <summary>
    ''' La table porte-t-elle DebutPeriode et FinPeriode ?
    '''
    ''' LA QUESTION SE POSE PARCE QUE LA TABLE EXISTE DÉJÀ EN PRODUCTION. Son CREATE TABLE est
    ''' ignoré au profit de celle qui est là, et les deux colonnes ne sont posées que par
    ''' l'ALTER de rattrapage du script d'installation. Tant qu'il n'a pas été joué,
    ''' l'application doit continuer à travailler sans la période — elle ne doit pas refuser
    ''' d'enregistrer une journée parce qu'un script d'évolution est en retard.
    '''
    ''' UN ÉCHEC DE LECTURE N'EST PAS MIS EN CACHE : base injoignable au démarrage puis
    ''' joignable ensuite, une réponse « non » retenue priverait de période toute la session.
    ''' </summary>
    Public Shared Function PeriodeDisponible() As Boolean

        If _periodeDisponible.HasValue Then Return _periodeDisponible.Value

        Try
            Using connexion As SqlConnection = WURepository.CreerConnexion()
                connexion.Open()
                _periodeDisponible = WURepository.ColonneExiste(connexion, Nothing, TABLE, "DebutPeriode")
            End Using

        Catch ex As SqlException
            Return False

        Catch ex As InvalidOperationException
            Return False
        End Try

        Return _periodeDisponible.Value
    End Function

    ''' <summary>
    ''' Les colonnes à lire : les fixes, plus celles de période quand la base les porte.
    ''' Les demander à une base qui ne les a pas ferait échouer la lecture entière.
    ''' </summary>
    Private Shared Function ColonnesLues() As String

        If PeriodeDisponible() Then Return COLONNES & ", DebutPeriode, FinPeriode"
        Return COLONNES
    End Function

#Region "Écriture"

    ''' <summary>
    ''' Écrit l'en-tête d'une journée, DANS LA TRANSACTION DE L'APPELANT.
    '''
    ''' La table absente n'interrompt PAS la comptabilisation : c'est l'opération du jour, elle
    ''' ne s'arrête pas parce qu'un script d'évolution n'a pas été joué. La fonction retourne
    ''' alors faux, et l'appelant le signale sans rien annuler.
    ''' </summary>
    Public Shared Function Enregistrer(traitement As TraitementJourneeWU,
                                       connexion As SqlConnection,
                                       transaction As SqlTransaction) As Boolean

        If traitement Is Nothing Then Return False

        If Not WURepository.ColonneExiste(connexion, transaction, TABLE, "DateActivite") Then
            Return False
        End If

        Const suppression As String = "DELETE FROM " & TABLE & " WHERE DateActivite = @jour"

        ' LA PÉRIODE N'EST ÉCRITE QUE SI LA BASE LA PORTE. La table existe en production : son
        ' CREATE TABLE est ignoré, et les deux colonnes ne viennent que de l'ALTER de
        ' rattrapage du script d'installation. Tant qu'il n'a pas été joué, la journée
        ' s'enregistre sans sa période plutôt que pas du tout — c'est l'opération du jour,
        ' elle ne s'arrête pas parce qu'un script d'évolution est en retard.
        Dim avecPeriode As Boolean =
            WURepository.ColonneExiste(connexion, transaction, TABLE, "DebutPeriode")

        Dim insertion As String =
            "INSERT INTO " & TABLE & " (DateActivite, DateValeur, NumeroLot, " &
            "FichierActivite, EmpreinteActivite, FichierReglement, EmpreinteReglement, " &
            "NombrePdv, NombreSousAgents, NombreAgences, NombreEcartes, " &
            "NombreEnvois, NombrePaiements, NombreAnnulations, " &
            "TotalDebit, TotalCredit, EcartArrondi, CompteEcart, " &
            "ComptabilisePar, DateComptabilisation" &
            If(avecPeriode, ", DebutPeriode, FinPeriode", String.Empty) & ") " &
            "VALUES (@jour, @valeur, @lot, @fichierA, @empreinteA, @fichierR, @empreinteR, " &
            "@pdv, @sousAgents, @agences, @ecartes, @envois, @paiements, @annulations, " &
            "@debit, @credit, @ecart, @compteEcart, @auteur, GETDATE()" &
            If(avecPeriode, ", @debutPeriode, @finPeriode", String.Empty) & ")"

        ' Le visa n'est PAS repris : recomptabiliser une journée l'efface, et c'est voulu.
        Using commande As New SqlCommand(suppression, connexion, transaction)
            commande.Parameters.Add("@jour", SqlDbType.Date).Value = traitement.DateActivite.Date
            commande.ExecuteNonQuery()
        End Using

        Using commande As New SqlCommand(insertion, connexion, transaction)

            commande.Parameters.Add("@jour", SqlDbType.Date).Value = traitement.DateActivite.Date

            commande.Parameters.Add("@valeur", SqlDbType.Date).Value =
                If(traitement.DateValeur.HasValue,
                   CType(traitement.DateValeur.Value.Date, Object), DBNull.Value)

            commande.Parameters.Add("@lot", SqlDbType.NVarChar, 10).Value = Texte(traitement.NumeroLot)
            commande.Parameters.Add("@fichierA", SqlDbType.NVarChar, 255).Value = Texte(traitement.FichierActivite)
            commande.Parameters.Add("@empreinteA", SqlDbType.NVarChar, 64).Value = Texte(traitement.EmpreinteActivite)
            commande.Parameters.Add("@fichierR", SqlDbType.NVarChar, 255).Value = Texte(traitement.FichierReglement)
            commande.Parameters.Add("@empreinteR", SqlDbType.NVarChar, 64).Value = Texte(traitement.EmpreinteReglement)

            commande.Parameters.Add("@pdv", SqlDbType.Int).Value = traitement.NombrePdv
            commande.Parameters.Add("@sousAgents", SqlDbType.Int).Value = traitement.NombreSousAgents
            commande.Parameters.Add("@agences", SqlDbType.Int).Value = traitement.NombreAgences
            commande.Parameters.Add("@ecartes", SqlDbType.Int).Value = traitement.NombreEcartes

            commande.Parameters.Add("@envois", SqlDbType.Int).Value = traitement.NombreEnvois
            commande.Parameters.Add("@paiements", SqlDbType.Int).Value = traitement.NombrePaiements
            commande.Parameters.Add("@annulations", SqlDbType.Int).Value = traitement.NombreAnnulations

            commande.Parameters.Add("@debit", SqlDbType.BigInt).Value = traitement.TotalDebit
            commande.Parameters.Add("@credit", SqlDbType.BigInt).Value = traitement.TotalCredit
            commande.Parameters.Add("@ecart", SqlDbType.BigInt).Value = traitement.EcartArrondi
            commande.Parameters.Add("@compteEcart", SqlDbType.NVarChar, 50).Value = Texte(traitement.CompteEcart)

            commande.Parameters.Add("@auteur", SqlDbType.NVarChar, 50).Value = SessionWU.Auteur

            If avecPeriode Then

                commande.Parameters.Add("@debutPeriode", SqlDbType.Date).Value =
                    If(traitement.DebutPeriode.HasValue,
                       CType(traitement.DebutPeriode.Value.Date, Object), DBNull.Value)

                commande.Parameters.Add("@finPeriode", SqlDbType.Date).Value =
                    If(traitement.FinPeriode.HasValue,
                       CType(traitement.FinPeriode.Value.Date, Object), DBNull.Value)
            End If

            commande.ExecuteNonQuery()
        End Using

        Return True
    End Function

    Private Shared Function Texte(valeur As String) As String
        Return If(valeur, String.Empty)
    End Function

#End Region

#Region "Le visa"

    ''' <summary>
    ''' Marque une journée visée.
    '''
    ''' Trois contrôles, dont deux sont redoublés dans la base : la fonction d'authorizer,
    ''' l'interdiction de viser son propre traitement, et le fait que la journée n'ait pas déjà
    ''' été visée. Le troisième se vérifie DANS la requête — deux supérieurs pourraient sinon
    ''' viser la même journée en même temps, et le second effacerait le premier.
    ''' </summary>
    Public Shared Function Viser(jour As Date, commentaire As String,
                                 ByRef messageErreur As String) As Boolean

        messageErreur = String.Empty

        If Not SessionWU.PeutAutoriserLesPointsDeVente Then
            messageErreur = "Seul un utilisateur ayant la fonction « authorizer » peut viser " &
                            "une journée de compensation."
            Return False
        End If

        ' La journée est relue d'abord : le message doit dire POURQUOI le visa est refusé,
        ' et non se contenter d'un « aucune ligne modifiée ».
        Dim traitement As TraitementJourneeWU = Charger(jour, messageErreur)
        If messageErreur.Length > 0 Then Return False

        If traitement Is Nothing OrElse Not traitement.Enregistre Then
            messageErreur = $"La journée du {jour:dd/MM/yyyy} n'a pas d'en-tête de traitement." &
                            Environment.NewLine &
                            "Elle a été comptabilisée avant la mise en service du bordereau, " &
                            "ou elle n'a jamais été comptabilisée. Elle ne peut pas être visée."
            Return False
        End If

        If traitement.EstVisee Then
            messageErreur = $"Cette journée a déjà été visée : {traitement.LibelleVisa}."
            Return False
        End If

        If String.Equals(traitement.ComptabilisePar, SessionWU.Auteur, StringComparison.OrdinalIgnoreCase) Then
            messageErreur = "Vous avez comptabilisé cette journée : vous ne pouvez pas la viser " &
                            "vous-même." & Environment.NewLine & Environment.NewLine &
                            "C'est tout l'objet du visa — un seul agent ne doit pas pouvoir " &
                            "arrêter seul une journée qui engage la comptabilité de la banque."
            Return False
        End If

        Const requete As String =
            "UPDATE " & TABLE & " SET VisePar = @viseur, DateVisa = GETDATE(), " &
            "CommentaireVisa = @commentaire " &
            "WHERE DateActivite = @jour AND VisePar IS NULL"

        Try
            Using connexion As SqlConnection = WURepository.CreerConnexion()
                connexion.Open()

                Using commande As New SqlCommand(requete, connexion)

                    commande.Parameters.Add("@jour", SqlDbType.Date).Value = jour.Date
                    commande.Parameters.Add("@viseur", SqlDbType.NVarChar, 50).Value = SessionWU.Auteur
                    commande.Parameters.Add("@commentaire", SqlDbType.NVarChar, 500).Value =
                        Texte(commentaire).Trim()

                    If commande.ExecuteNonQuery() = 0 Then
                        messageErreur = "Cette journée vient d'être visée par quelqu'un d'autre." &
                                        Environment.NewLine & "Actualisez l'écran."
                        Return False
                    End If
                End Using
            End Using

        Catch ex As SqlException
            messageErreur = If(ex.Number = ERREUR_TABLE_ABSENTE,
                               MESSAGE_TABLE_ABSENTE,
                               $"Visa impossible : {ex.Message}")
            Return False

        Catch ex As InvalidOperationException
            messageErreur = $"Connexion SQL Server indisponible : {ex.Message}"
            Return False
        End Try

        Return True
    End Function

#End Region

#Region "Lecture"

    ''' <summary>
    ''' L'en-tête d'une journée, ou un en-tête NON ENREGISTRÉ si la table ne le contient pas.
    '''
    ''' La distinction compte : un en-tête absent n'est pas une erreur mais une journée
    ''' antérieure à la mise en service, et le bordereau doit pouvoir le dire plutôt que de
    ''' refuser de s'afficher.
    ''' </summary>
    Public Shared Function Charger(jour As Date, ByRef messageErreur As String) As TraitementJourneeWU

        messageErreur = String.Empty

        Dim requete As String =
            "SELECT " & ColonnesLues() & " FROM " & TABLE & " WHERE DateActivite = @jour"

        Try
            Using connexion As SqlConnection = WURepository.CreerConnexion()
                connexion.Open()

                If Not WURepository.ColonneExiste(connexion, Nothing, TABLE, "DateActivite") Then
                    Return New TraitementJourneeWU() With {.DateActivite = jour.Date}
                End If

                Using commande As New SqlCommand(requete, connexion)
                    commande.Parameters.Add("@jour", SqlDbType.Date).Value = jour.Date

                    Using lecteur As SqlDataReader = commande.ExecuteReader()

                        If Not lecteur.Read() Then
                            Return New TraitementJourneeWU() With {.DateActivite = jour.Date}
                        End If

                        Return Construire(lecteur)
                    End Using
                End Using
            End Using

        Catch ex As SqlException
            messageErreur = If(ex.Number = ERREUR_TABLE_ABSENTE,
                               MESSAGE_TABLE_ABSENTE,
                               $"Lecture de l'en-tête de traitement impossible : {ex.Message}")
            Return Nothing

        Catch ex As InvalidOperationException
            messageErreur = $"Connexion SQL Server indisponible : {ex.Message}"
            Return Nothing
        End Try
    End Function

    ''' <summary>Les journées comptabilisées et non encore visées, la plus ancienne d'abord.</summary>
    Public Shared Function ListerEnAttenteDeVisa(ByRef messageErreur As String) As List(Of TraitementJourneeWU)

        messageErreur = String.Empty
        Dim resultat As New List(Of TraitementJourneeWU)()

        Dim requete As String =
            "SELECT " & ColonnesLues() & " FROM " & TABLE & " WHERE VisePar IS NULL ORDER BY DateActivite"

        Try
            Using connexion As SqlConnection = WURepository.CreerConnexion()
                connexion.Open()

                If Not WURepository.ColonneExiste(connexion, Nothing, TABLE, "DateActivite") Then
                    Return resultat
                End If

                Using commande As New SqlCommand(requete, connexion)
                    Using lecteur As SqlDataReader = commande.ExecuteReader()
                        While lecteur.Read()
                            resultat.Add(Construire(lecteur))
                        End While
                    End Using
                End Using
            End Using

        Catch ex As SqlException
            messageErreur = If(ex.Number = ERREUR_TABLE_ABSENTE,
                               MESSAGE_TABLE_ABSENTE,
                               $"Lecture des journées à viser impossible : {ex.Message}")

        Catch ex As InvalidOperationException
            messageErreur = $"Connexion SQL Server indisponible : {ex.Message}"
        End Try

        Return resultat
    End Function

    ''' <summary>
    ''' Les journées déjà comptabilisées dont la période CHEVAUCHE celle qu'on s'apprête à
    ''' comptabiliser. Liste vide s'il n'y en a aucune.
    '''
    ''' CE CONTRÔLE N'EXISTAIT PAS, ET L'ÉCRAN LE PROMETTAIT DÉJÀ. Avant de comptabiliser, il
    ''' avertit : « ne chargez pas ensuite un rapport d'une journée déjà comprise dans cette
    ''' période, elle serait comptabilisée deux fois ». Mais il ne cherchait qu'une
    ''' comptabilisation à la DATE D'ACTIVITÉ. Recharger le 24/09 seul après une semaine
    ''' comptabilisée du 24 au 30 ne trouvait rien, et passait sans un mot : les six journées
    ''' étaient enregistrées sous la seule date du premier jour, et les cinq autres restaient
    ''' invisibles à toute recherche.
    '''
    ''' DEUX PÉRIODES SE CHEVAUCHENT quand chacune commence avant que l'autre ne finisse. Les
    ''' bornes sont incluses : une semaine qui finit le 30 et une autre qui commence le 30
    ''' partagent bien une journée.
    '''
    ''' UNE JOURNÉE ENREGISTRÉE AVANT CETTE VERSION n'a pas de période : ISNULL la ramène à sa
    ''' date d'activité, c'est-à-dire à une période d'un seul jour. Le contrôle reste donc
    ''' exact pour les journées simples, et seulement aveugle aux anciennes semaines — qu'il
    ''' n'aurait de toute façon pas pu voir.
    ''' </summary>
    Public Shared Function PeriodesQuiChevauchent(debut As Date, fin As Date,
                                                  ByRef messageErreur As String) As List(Of TraitementJourneeWU)

        messageErreur = String.Empty
        Dim resultat As New List(Of TraitementJourneeWU)()

        Dim premier As Date = If(debut <= fin, debut.Date, fin.Date)
        Dim dernier As Date = If(debut <= fin, fin.Date, debut.Date)

        Dim filtre As String =
            If(PeriodeDisponible(),
               "ISNULL(FinPeriode, DateActivite) >= @debut AND ISNULL(DebutPeriode, DateActivite) <= @fin",
               "DateActivite >= @debut AND DateActivite <= @fin")

        Dim requete As String =
            "SELECT " & ColonnesLues() & " FROM " & TABLE & " WHERE " & filtre & " ORDER BY DateActivite"

        Try
            Using connexion As SqlConnection = WURepository.CreerConnexion()
                connexion.Open()

                If Not WURepository.ColonneExiste(connexion, Nothing, TABLE, "DateActivite") Then
                    Return resultat
                End If

                Using commande As New SqlCommand(requete, connexion)

                    commande.Parameters.Add("@debut", SqlDbType.Date).Value = premier
                    commande.Parameters.Add("@fin", SqlDbType.Date).Value = dernier

                    Using lecteur As SqlDataReader = commande.ExecuteReader()
                        While lecteur.Read()
                            resultat.Add(Construire(lecteur))
                        End While
                    End Using
                End Using
            End Using

        Catch ex As SqlException
            messageErreur = If(ex.Number = ERREUR_TABLE_ABSENTE,
                               MESSAGE_TABLE_ABSENTE,
                               $"Recherche des journées déjà comptabilisées impossible : {ex.Message}")

        Catch ex As InvalidOperationException
            messageErreur = $"Connexion SQL Server indisponible : {ex.Message}"
        End Try

        Return resultat
    End Function

    Private Shared Function Construire(lecteur As SqlDataReader) As TraitementJourneeWU

        Return New TraitementJourneeWU() With {
            .DateActivite = Convert.ToDateTime(lecteur("DateActivite"), CultureInfo.InvariantCulture),
            .DateValeur = LireDate(lecteur, "DateValeur"),
            .NumeroLot = LireChaine(lecteur, "NumeroLot"),
            .FichierActivite = LireChaine(lecteur, "FichierActivite"),
            .EmpreinteActivite = LireChaine(lecteur, "EmpreinteActivite"),
            .FichierReglement = LireChaine(lecteur, "FichierReglement"),
            .EmpreinteReglement = LireChaine(lecteur, "EmpreinteReglement"),
            .NombrePdv = LireEntier(lecteur, "NombrePdv"),
            .NombreSousAgents = LireEntier(lecteur, "NombreSousAgents"),
            .NombreAgences = LireEntier(lecteur, "NombreAgences"),
            .NombreEcartes = LireEntier(lecteur, "NombreEcartes"),
            .NombreEnvois = LireEntier(lecteur, "NombreEnvois"),
            .NombrePaiements = LireEntier(lecteur, "NombrePaiements"),
            .NombreAnnulations = LireEntier(lecteur, "NombreAnnulations"),
            .TotalDebit = LireEntierLong(lecteur, "TotalDebit"),
            .TotalCredit = LireEntierLong(lecteur, "TotalCredit"),
            .EcartArrondi = LireEntierLong(lecteur, "EcartArrondi"),
            .CompteEcart = LireChaine(lecteur, "CompteEcart"),
            .ComptabilisePar = LireChaine(lecteur, "ComptabilisePar"),
            .DateComptabilisation = LireDate(lecteur, "DateComptabilisation"),
            .VisePar = LireChaine(lecteur, "VisePar"),
            .DateVisa = LireDate(lecteur, "DateVisa"),
            .CommentaireVisa = LireChaine(lecteur, "CommentaireVisa"),
            .DebutPeriode = LireDateSiPresente(lecteur, "DebutPeriode"),
            .FinPeriode = LireDateSiPresente(lecteur, "FinPeriode"),
            .Enregistre = True
        }
    End Function

#End Region

#Region "Utilitaires de lecture"

    ''' <summary>
    ''' Une date que le lecteur ne porte pas forcément : Nothing si la colonne est absente du
    ''' jeu de résultats, et non une exception.
    '''
    ''' GetOrdinal LÈVE une IndexOutOfRangeException sur une colonne inconnue. Comme la liste
    ''' des colonnes lues dépend de ce que la base porte, demander la période sans vérifier
    ''' ferait échouer la lecture de toutes les journées sur une base en retard d'un script.
    ''' </summary>
    Private Shared Function LireDateSiPresente(lecteur As SqlDataReader, colonne As String) As Date?

        If Not ColonneDansLecteur(lecteur, colonne) Then Return Nothing
        Return LireDate(lecteur, colonne)
    End Function

    ''' <summary>Vrai si le jeu de résultats porte cette colonne, quelle que soit la casse.</summary>
    Private Shared Function ColonneDansLecteur(lecteur As SqlDataReader, colonne As String) As Boolean

        For index As Integer = 0 To lecteur.FieldCount - 1
            If String.Equals(lecteur.GetName(index), colonne, StringComparison.OrdinalIgnoreCase) Then
                Return True
            End If
        Next

        Return False
    End Function

    Private Shared Function LireChaine(lecteur As SqlDataReader, colonne As String) As String
        Dim index As Integer = lecteur.GetOrdinal(colonne)
        If lecteur.IsDBNull(index) Then Return String.Empty
        Return lecteur.GetValue(index).ToString().Trim()
    End Function

    Private Shared Function LireEntier(lecteur As SqlDataReader, colonne As String) As Integer
        Dim index As Integer = lecteur.GetOrdinal(colonne)
        If lecteur.IsDBNull(index) Then Return 0
        Return Convert.ToInt32(lecteur.GetValue(index), CultureInfo.InvariantCulture)
    End Function

    Private Shared Function LireEntierLong(lecteur As SqlDataReader, colonne As String) As Long
        Dim index As Integer = lecteur.GetOrdinal(colonne)
        If lecteur.IsDBNull(index) Then Return 0L
        Return Convert.ToInt64(lecteur.GetValue(index), CultureInfo.InvariantCulture)
    End Function

    Private Shared Function LireDate(lecteur As SqlDataReader, colonne As String) As Date?
        Dim index As Integer = lecteur.GetOrdinal(colonne)
        If lecteur.IsDBNull(index) Then Return Nothing
        Return Convert.ToDateTime(lecteur.GetValue(index), CultureInfo.InvariantCulture)
    End Function

#End Region

End Class
