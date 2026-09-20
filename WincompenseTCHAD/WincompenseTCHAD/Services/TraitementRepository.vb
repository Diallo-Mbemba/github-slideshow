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
        "Exécutez le script Scripts\17_BordereauJournee.sql : il la crée." & vbCrLf &
        "Tant qu'elle est absente, le bordereau de fin de journée se reconstitue depuis " &
        "l'historique et la pièce, mais sans le nom des rapports Western Union — et aucune " &
        "journée ne peut être visée."

    Private Const COLONNES As String =
        "DateActivite, DateValeur, NumeroLot, FichierActivite, EmpreinteActivite, " &
        "FichierReglement, EmpreinteReglement, NombrePdv, NombreSousAgents, NombreAgences, " &
        "NombreEcartes, NombreEnvois, NombrePaiements, NombreAnnulations, " &
        "TotalDebit, TotalCredit, EcartArrondi, CompteEcart, " &
        "ComptabilisePar, DateComptabilisation, VisePar, DateVisa, CommentaireVisa"

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

        Const insertion As String =
            "INSERT INTO " & TABLE & " (DateActivite, DateValeur, NumeroLot, " &
            "FichierActivite, EmpreinteActivite, FichierReglement, EmpreinteReglement, " &
            "NombrePdv, NombreSousAgents, NombreAgences, NombreEcartes, " &
            "NombreEnvois, NombrePaiements, NombreAnnulations, " &
            "TotalDebit, TotalCredit, EcartArrondi, CompteEcart, " &
            "ComptabilisePar, DateComptabilisation) " &
            "VALUES (@jour, @valeur, @lot, @fichierA, @empreinteA, @fichierR, @empreinteR, " &
            "@pdv, @sousAgents, @agences, @ecartes, @envois, @paiements, @annulations, " &
            "@debit, @credit, @ecart, @compteEcart, @auteur, GETDATE())"

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
            "SELECT " & COLONNES & " FROM " & TABLE & " WHERE DateActivite = @jour"

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
            "SELECT " & COLONNES & " FROM " & TABLE & " WHERE VisePar IS NULL ORDER BY DateActivite"

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
            .Enregistre = True
        }
    End Function

#End Region

#Region "Utilitaires de lecture"

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
