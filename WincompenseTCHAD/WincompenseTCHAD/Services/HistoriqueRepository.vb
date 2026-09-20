Option Strict On
Option Explicit On

Imports System.Data
Imports System.Data.SqlClient

''' <summary>
''' Historique des journées comptabilisées (table T_HistoriqueWU) : écriture au moment de la
''' génération de la pièce comptable, lecture pour les rapports d'activité sur une période.
'''
''' L'historique restitue ce qui a RÉELLEMENT été comptabilisé, et non un recalcul a posteriori.
''' C'est pourquoi il recopie l'identification du point de vente telle qu'elle était ce jour-là
''' plutôt que de la rattacher au paramétrage courant.
'''
''' Comme les autres dépôts, aucune exception SQL ne remonte à l'interface : chaque fonction
''' retourne un booléen de réussite et pose un message explicite, en français.
''' </summary>
Public NotInheritable Class HistoriqueRepository

    Private Sub New()
    End Sub

    Private Const TABLE_HISTORIQUE As String = "T_HistoriqueWU"

    ''' <summary>Code d'erreur SQL Server signalant une table absente (« Invalid object name »).</summary>
    Private Const ERREUR_TABLE_ABSENTE As Integer = 208

    ''' <summary>Message posé lorsque la table n'existe pas : le script d'installation n'a pas été joué.</summary>
    Public Const MESSAGE_TABLE_ABSENTE As String =
        "La table T_HistoriqueWU n'existe pas encore dans la base." & vbCrLf & vbCrLf &
        "Exécutez le script Scripts\05_HistoriqueWU.sql : il la crée." & vbCrLf &
        "Tant qu'elle est absente, les journées comptabilisées ne sont pas historisées et les " &
        "rapports d'activité restent vides."

    ''' <summary>Message posé lorsque la table du détail des transactions n'existe pas.</summary>
    Public Const MESSAGE_TABLE_MTCN_ABSENTE As String =
        "La table T_HistoriqueMTCN n'existe pas encore dans la base." & vbCrLf & vbCrLf &
        "Exécutez le script Scripts\06_HistoriqueMTCN.sql : il la crée." & vbCrLf &
        "Tant qu'elle est absente, le détail des MTCN n'est pas conservé."

#Region "Une journée déjà comptabilisée"

    ''' <summary>
    ''' Ce que l'application sait d'une journée déjà comptabilisée : de quoi avertir avant
    ''' de la remplacer.
    ''' </summary>
    Public NotInheritable Class Comptabilisation

        Public Property DateActivite As Date
        Public Property NombrePdv As Integer = 0
        Public Property DateComptabilisation As Date?
        Public Property ComptabilisePar As String = String.Empty

        ''' <summary>
        ''' La phrase d'avertissement. Elle nomme la personne et l'heure : « déjà
        ''' comptabilisée » ne suffit pas à décider, « comptabilisée il y a dix minutes par
        ''' votre collègue » si.
        ''' </summary>
        Public ReadOnly Property Avertissement As String
            Get
                Dim phrase As String =
                    $"La journée du {DateActivite:dd/MM/yyyy} a déjà été comptabilisée"

                If DateComptabilisation.HasValue Then
                    phrase &= $" le {DateComptabilisation.Value:dd/MM/yyyy à HH:mm}"
                End If

                If ComptabilisePar.Length > 0 Then phrase &= $" par {ComptabilisePar}"

                Return phrase & $" ({NombrePdv} point(s) de vente)."
            End Get
        End Property
    End Class

    ''' <summary>
    ''' La comptabilisation en vigueur pour une journée, ou Nothing si elle n'a jamais été
    ''' comptabilisée — ou si elle a été annulée depuis, ce qui revient au même : il n'y a
    ''' alors plus rien à remplacer.
    '''
    ''' Une table absente ou une base injoignable ne renvoient rien : mieux vaut ne pas
    ''' avertir que bloquer une comptabilisation pour un avertissement.
    ''' </summary>
    Public Shared Function ComptabilisationExistante(jour As Date) As Comptabilisation

        Const lecture As String =
            "SELECT  pdv     = COUNT(*)," &
            "        quand   = MAX(DateComptabilisation)," &
            "        qui     = MIN(ComptabilisePar) " &
            "FROM    " & TABLE_HISTORIQUE & " " &
            "WHERE   DateActivite = @jour"

        Try
            Using connexion As SqlConnection = WURepository.CreerConnexion()
                connexion.Open()

                Using commande As New SqlCommand(lecture, connexion)
                    commande.Parameters.Add("@jour", SqlDbType.Date).Value = jour.Date

                    Using lecteur As SqlDataReader = commande.ExecuteReader()

                        If Not lecteur.Read() Then Return Nothing

                        Dim nombre As Integer = Convert.ToInt32(lecteur.GetValue(0),
                                                                Globalization.CultureInfo.InvariantCulture)
                        If nombre = 0 Then Return Nothing

                        Return New Comptabilisation() With {
                            .DateActivite = jour.Date,
                            .NombrePdv = nombre,
                            .DateComptabilisation = If(lecteur.IsDBNull(1), CType(Nothing, Date?),
                                                       CType(lecteur.GetDateTime(1), Date?)),
                            .ComptabilisePar = If(lecteur.IsDBNull(2), String.Empty,
                                                  lecteur.GetValue(2).ToString().Trim())
                        }
                    End Using
                End Using
            End Using

        Catch ex As SqlException
            Return Nothing
        Catch ex As InvalidOperationException
            Return Nothing
        End Try
    End Function

#End Region

#Region "Écriture"

    ''' <summary>
    ''' Enregistre le résultat d'une journée comptabilisée : une ligne par point de vente.
    '''
    ''' L'écriture est RÉPÉTABLE : les lignes existantes de la journée sont d'abord supprimées.
    ''' Une journée regénérée — après correction d'un paramétrage, par exemple — remplace donc
    ''' intégralement la précédente, sans jamais produire de doublon ni de cumul.
    '''
    ''' Suppression et insertion se font dans une même transaction : en cas d'incident, la
    ''' journée reste dans son état antérieur plutôt qu'amputée de ses lignes.
    ''' </summary>
    ''' <param name="jour">Journée d'activité comptabilisée.</param>
    ''' <param name="calculs">Résultat du calcul, un élément par Account.</param>
    ''' <param name="nombreEnregistrees">Nombre de lignes effectivement écrites.</param>
    ''' <param name="transactions">
    ''' Détail des transactions de la journée, écrit dans T_HistoriqueMTCN au sein de la MÊME
    ''' transaction : les deux tables ne peuvent pas diverger. Peut être Nothing.
    ''' </param>
    ''' <param name="dtPiece">
    ''' Pièce comptable produite pour cette journée, conservée dans T_PieceWU au sein de la MÊME
    ''' transaction — pour la même raison, et pour une de plus : la pièce est un justificatif,
    ''' et un justificatif qui ne correspondrait pas à l'historique de sa propre journée serait
    ''' pire que pas de justificatif du tout. Peut être Nothing.
    ''' </param>
    Public Shared Function EnregistrerJournee(jour As Date, calculs As IEnumerable(Of CalculWU),
                                              transactions As IEnumerable(Of TransactionWU),
                                              dtPiece As DataTable,
                                              ByRef nombreEnregistrees As Integer,
                                              ByRef nombreTransactions As Integer,
                                              ByRef messageErreur As String) As Boolean

        messageErreur = String.Empty
        nombreEnregistrees = 0
        nombreTransactions = 0

        If calculs Is Nothing Then
            messageErreur = "Aucun calcul à historiser."
            Return False
        End If

        Const suppression As String = "DELETE FROM " & TABLE_HISTORIQUE & " WHERE DateActivite = @jour"

        Const insertion As String =
            "INSERT INTO " & TABLE_HISTORIQUE & " (DateActivite, Account, Designation, " &
            "GroupeStatistique, TypePdv, NombreEnvois, NombrePaiements, NombreAnnulations, " &
            "PrincipalEnvoi, ChargeEnvoi, Taxes, PrincipalPaye, " &
            "CommissionEnvoi, CommissionPaiement, CommissionTransfert, " &
            "TVA, TTAEnvoi, TTAReception, TaxeEnvoi, DateComptabilisation, ComptabilisePar) " &
            "VALUES (@jour, @account, @designation, @groupe, @type, @nbEnvois, @nbPaiements, " &
            "@nbAnnulations, @principalEnvoi, @chargeEnvoi, @taxes, @principalPaye, " &
            "@comEnvoi, @comPaiement, @comTransfert, @tva, @ttaEnvoi, @ttaReception, @taxeEnvoi, " &
            "GETDATE(), @auteur)"

        Try
            Using connexion As SqlConnection = WURepository.CreerConnexion()
                connexion.Open()

                Using transaction As SqlTransaction = connexion.BeginTransaction()
                    Try
                        Using commande As New SqlCommand(suppression, connexion, transaction)
                            commande.Parameters.Add("@jour", SqlDbType.Date).Value = jour.Date
                            commande.ExecuteNonQuery()
                        End Using

                        For Each calc As CalculWU In calculs

                            ' Un Account sans opération de la journée n'a rien à historiser.
                            If calc Is Nothing OrElse String.IsNullOrWhiteSpace(calc.Account) Then Continue For

                            Dim ligne As LigneHistoriqueWU = LigneHistoriqueWU.DepuisCalcul(calc, jour)

                            Using commande As New SqlCommand(insertion, connexion, transaction)
                                AjouterParametres(commande, ligne)
                                commande.ExecuteNonQuery()
                            End Using

                            nombreEnregistrees += 1
                        Next

                        ' Identification des points de vente, pour compléter le détail des
                        ' transactions : le rapport d'activité ne la porte pas.
                        Dim identification As New Dictionary(Of String, CalculWU)(StringComparer.OrdinalIgnoreCase)
                        For Each calc As CalculWU In calculs
                            If calc IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(calc.Account) Then
                                identification(calc.Account) = calc
                            End If
                        Next

                        nombreTransactions = EcrireTransactions(connexion, transaction, jour,
                                                                transactions, identification)

                        ' La pièce est conservée ici, et non dans un appel séparé : l'historique
                        ' et le justificatif d'une même journée ne doivent jamais diverger.
                        PieceRepository.Enregistrer(jour, dtPiece, connexion, transaction)

                        transaction.Commit()

                    Catch
                        Try
                            transaction.Rollback()
                        Catch
                            ' La connexion est peut-être déjà tombée : l'erreur d'origine prime.
                        End Try
                        Throw
                    End Try
                End Using
            End Using

        Catch ex As SqlException
            nombreEnregistrees = 0
            nombreTransactions = 0
            messageErreur = If(ex.Number = ERREUR_TABLE_ABSENTE,
                               MESSAGE_TABLE_ABSENTE & vbCrLf & vbCrLf &
                               "Si seule T_HistoriqueMTCN manque, exécutez aussi " &
                               "Scripts\06_HistoriqueMTCN.sql ; si c'est T_PieceWU, " &
                               "Scripts\13_PiecesComptables.sql.",
                               $"Historisation de la journée du {jour:dd/MM/yyyy} impossible : {ex.Message}")
            Return False

        Catch ex As InvalidOperationException
            nombreEnregistrees = 0
            nombreTransactions = 0
            messageErreur = $"Connexion SQL Server indisponible : {ex.Message}"
            Return False
        End Try

        Return True
    End Function

    Private Shared Sub AjouterParametres(commande As SqlCommand, ligne As LigneHistoriqueWU)

        commande.Parameters.Add("@jour", SqlDbType.Date).Value = ligne.DateActivite
        commande.Parameters.Add("@account", SqlDbType.NVarChar, 255).Value = ligne.Account
        commande.Parameters.Add("@designation", SqlDbType.NVarChar, 255).Value = If(ligne.Designation, String.Empty)
        commande.Parameters.Add("@groupe", SqlDbType.NVarChar, 255).Value = If(ligne.GroupeStatistique, String.Empty)
        commande.Parameters.Add("@type", SqlDbType.NVarChar, 20).Value = If(ligne.TypePdv, String.Empty)

        commande.Parameters.Add("@nbEnvois", SqlDbType.Int).Value = ligne.NombreEnvois
        commande.Parameters.Add("@nbPaiements", SqlDbType.Int).Value = ligne.NombrePaiements
        commande.Parameters.Add("@nbAnnulations", SqlDbType.Int).Value = ligne.NombreAnnulations

        AjouterMontant(commande, "@principalEnvoi", ligne.PrincipalEnvoi)
        AjouterMontant(commande, "@chargeEnvoi", ligne.ChargeEnvoi)
        AjouterMontant(commande, "@taxes", ligne.Taxes)
        AjouterMontant(commande, "@principalPaye", ligne.PrincipalPaye)
        AjouterMontant(commande, "@comEnvoi", ligne.CommissionEnvoi)
        AjouterMontant(commande, "@comPaiement", ligne.CommissionPaiement)
        AjouterMontant(commande, "@comTransfert", ligne.CommissionTransfert)
        AjouterMontant(commande, "@tva", ligne.TVA)
        AjouterMontant(commande, "@ttaEnvoi", ligne.TTAEnvoi)
        AjouterMontant(commande, "@ttaReception", ligne.TTAReception)
        AjouterMontant(commande, "@taxeEnvoi", ligne.TaxeEnvoi)

        commande.Parameters.Add("@auteur", SqlDbType.NVarChar, 50).Value = SessionWU.Auteur
    End Sub

    ''' <summary>
    ''' Ajoute un montant en cadrant explicitement précision et échelle sur DECIMAL(18,2) :
    ''' sans cela, ADO.NET déduit l'échelle de la valeur transmise et SQL Server peut arrondir
    ''' autrement qu'attendu.
    ''' </summary>
    Private Shared Sub AjouterMontant(commande As SqlCommand, nom As String, valeur As Decimal)

        Dim parametre As SqlParameter = commande.Parameters.Add(nom, SqlDbType.Decimal)
        parametre.Precision = 18
        parametre.Scale = 2
        parametre.Value = Decimal.Round(valeur, 2)
    End Sub

#End Region

#Region "Détail des transactions (T_HistoriqueMTCN)"

    Private Const TABLE_MTCN As String = "T_HistoriqueMTCN"

    ''' <summary>
    ''' Écrit le détail des transactions de la journée. Appelée DANS la transaction qui écrit
    ''' l'agrégat : les deux tables ne peuvent donc pas diverger.
    '''
    ''' L'identification du point de vente est complétée depuis les calculs de la journée, le
    ''' rapport d'activité ne la portant pas.
    ''' </summary>
    Private Shared Function EcrireTransactions(connexion As SqlConnection, transaction As SqlTransaction,
                                               jour As Date, transactions As IEnumerable(Of TransactionWU),
                                               identification As Dictionary(Of String, CalculWU)) As Integer

        Using commande As New SqlCommand("DELETE FROM " & TABLE_MTCN & " WHERE DateActivite = @jour",
                                         connexion, transaction)
            commande.Parameters.Add("@jour", SqlDbType.Date).Value = jour.Date
            commande.ExecuteNonQuery()
        End Using

        If transactions Is Nothing Then Return 0

        Const insertion As String =
            "INSERT INTO " & TABLE_MTCN & " (DateActivite, Account, MTCN, Sens, Statut, Montant, " &
            "Designation, GroupeStatistique, TypePdv, DateComptabilisation, ComptabilisePar) " &
            "VALUES (@jour, @account, @mtcn, @sens, @statut, @montant, @designation, @groupe, @type, " &
            "GETDATE(), @auteur)"

        Dim ecrites As Integer = 0

        For Each operation As TransactionWU In transactions

            If operation Is Nothing OrElse String.IsNullOrWhiteSpace(operation.MTCN) Then Continue For

            Dim calc As CalculWU = Nothing
            identification.TryGetValue(operation.Account, calc)

            Using commande As New SqlCommand(insertion, connexion, transaction)

                commande.Parameters.Add("@jour", SqlDbType.Date).Value = jour.Date
                commande.Parameters.Add("@account", SqlDbType.NVarChar, 255).Value = operation.Account
                commande.Parameters.Add("@mtcn", SqlDbType.NVarChar, 50).Value = operation.MTCN
                commande.Parameters.Add("@sens", SqlDbType.NVarChar, 10).Value = operation.Sens
                commande.Parameters.Add("@statut", SqlDbType.NVarChar, 10).Value = If(operation.Statut, String.Empty)

                Dim parametreMontant As SqlParameter = commande.Parameters.Add("@montant", SqlDbType.Decimal)
                parametreMontant.Precision = 18
                parametreMontant.Scale = 2
                parametreMontant.Value = Decimal.Round(operation.Montant, 2)

                commande.Parameters.Add("@designation", SqlDbType.NVarChar, 255).Value =
                    If(calc Is Nothing, String.Empty, If(calc.Designation, String.Empty))
                commande.Parameters.Add("@groupe", SqlDbType.NVarChar, 255).Value =
                    If(calc Is Nothing, String.Empty, If(calc.GroupeStatistique, String.Empty))
                commande.Parameters.Add("@type", SqlDbType.NVarChar, 20).Value =
                    If(calc Is Nothing, String.Empty, If(calc.TypePdv, String.Empty))
                commande.Parameters.Add("@auteur", SqlDbType.NVarChar, 50).Value = SessionWU.Auteur

                commande.ExecuteNonQuery()
            End Using

            ecrites += 1
        Next

        Return ecrites
    End Function

    ''' <summary>
    ''' Retourne le détail des transactions d'une période, éventuellement restreint à un groupe
    ''' statistique. Triées par date, point de vente puis MTCN.
    ''' </summary>
    ''' <param name="groupe">Groupe statistique, ou chaîne vide pour tous.</param>
    Public Shared Function ListerTransactions(debut As Date, fin As Date, groupe As String,
                                              ByRef messageErreur As String) As List(Of TransactionWU)

        messageErreur = String.Empty
        Dim resultat As New List(Of TransactionWU)

        Dim requete As String =
            "SELECT DateActivite, Account, MTCN, Sens, Statut, Montant, " &
            "Designation, GroupeStatistique, TypePdv " &
            "FROM " & TABLE_MTCN & " " &
            "WHERE DateActivite >= @debut AND DateActivite <= @fin"

        Dim filtreGroupe As String = If(groupe, String.Empty).Trim()
        If filtreGroupe.Length > 0 Then
            requete &= " AND LTRIM(RTRIM(ISNULL(GroupeStatistique, ''))) = @groupe"
        End If
        requete &= " ORDER BY DateActivite, Account, MTCN"

        Try
            Using connexion As SqlConnection = WURepository.CreerConnexion()
                connexion.Open()

                Using commande As New SqlCommand(requete, connexion)
                    commande.Parameters.Add("@debut", SqlDbType.Date).Value = debut.Date
                    commande.Parameters.Add("@fin", SqlDbType.Date).Value = fin.Date

                    If filtreGroupe.Length > 0 Then
                        commande.Parameters.Add("@groupe", SqlDbType.NVarChar, 255).Value = filtreGroupe
                    End If

                    Using lecteur As SqlDataReader = commande.ExecuteReader()
                        While lecteur.Read()
                            resultat.Add(New TransactionWU() With {
                                .DateActivite = Convert.ToDateTime(lecteur("DateActivite"), Globalization.CultureInfo.InvariantCulture),
                                .Account = LireChaine(lecteur, "Account"),
                                .MTCN = LireChaine(lecteur, "MTCN"),
                                .Sens = LireChaine(lecteur, "Sens"),
                                .Statut = LireChaine(lecteur, "Statut"),
                                .Montant = LireMontant(lecteur, "Montant"),
                                .Designation = LireChaine(lecteur, "Designation"),
                                .GroupeStatistique = LireChaine(lecteur, "GroupeStatistique"),
                                .TypePdv = LireChaine(lecteur, "TypePdv")
                            })
                        End While
                    End Using
                End Using
            End Using

        Catch ex As SqlException
            messageErreur = If(ex.Number = ERREUR_TABLE_ABSENTE,
                               MESSAGE_TABLE_MTCN_ABSENTE,
                               $"Lecture du détail des transactions impossible : {ex.Message}")
        Catch ex As InvalidOperationException
            messageErreur = $"Connexion SQL Server indisponible : {ex.Message}"
        End Try

        Return resultat
    End Function

#End Region

#Region "Lecture"

    ''' <summary>
    ''' Retourne les lignes d'historique d'une période, bornes comprises.
    '''
    ''' Les rapports agrègent ensuite ces lignes en mémoire — par jour, par point de vente,
    ''' par groupe — plutôt que d'interroger la base quatre fois. Un mois représente quelques
    ''' centaines de lignes, et cette source unique garantit que les quatre pages d'un même
    ''' rapport présentent exactement les mêmes totaux.
    ''' </summary>
    ''' <returns>Lignes triées par date puis par Account. Liste vide en cas d'erreur.</returns>
    Public Shared Function ListerPeriode(debut As Date, fin As Date, ByRef messageErreur As String) As List(Of LigneHistoriqueWU)

        messageErreur = String.Empty
        Dim resultat As New List(Of LigneHistoriqueWU)

        Const requete As String =
            "SELECT DateActivite, Account, Designation, GroupeStatistique, TypePdv, " &
            "NombreEnvois, NombrePaiements, NombreAnnulations, " &
            "PrincipalEnvoi, ChargeEnvoi, Taxes, PrincipalPaye, " &
            "CommissionEnvoi, CommissionPaiement, CommissionTransfert, " &
            "TVA, TTAEnvoi, TTAReception, TaxeEnvoi " &
            "FROM " & TABLE_HISTORIQUE & " " &
            "WHERE DateActivite >= @debut AND DateActivite <= @fin " &
            "ORDER BY DateActivite, Account"

        Try
            Using connexion As SqlConnection = WURepository.CreerConnexion()
                connexion.Open()

                Using commande As New SqlCommand(requete, connexion)
                    commande.Parameters.Add("@debut", SqlDbType.Date).Value = debut.Date
                    commande.Parameters.Add("@fin", SqlDbType.Date).Value = fin.Date

                    Using lecteur As SqlDataReader = commande.ExecuteReader()
                        While lecteur.Read()
                            resultat.Add(New LigneHistoriqueWU() With {
                                .DateActivite = Convert.ToDateTime(lecteur("DateActivite"), Globalization.CultureInfo.InvariantCulture),
                                .Account = LireChaine(lecteur, "Account"),
                                .Designation = LireChaine(lecteur, "Designation"),
                                .GroupeStatistique = LireChaine(lecteur, "GroupeStatistique"),
                                .TypePdv = LireChaine(lecteur, "TypePdv"),
                                .NombreEnvois = LireEntier(lecteur, "NombreEnvois"),
                                .NombrePaiements = LireEntier(lecteur, "NombrePaiements"),
                                .NombreAnnulations = LireEntier(lecteur, "NombreAnnulations"),
                                .PrincipalEnvoi = LireMontant(lecteur, "PrincipalEnvoi"),
                                .ChargeEnvoi = LireMontant(lecteur, "ChargeEnvoi"),
                                .Taxes = LireMontant(lecteur, "Taxes"),
                                .PrincipalPaye = LireMontant(lecteur, "PrincipalPaye"),
                                .CommissionEnvoi = LireMontant(lecteur, "CommissionEnvoi"),
                                .CommissionPaiement = LireMontant(lecteur, "CommissionPaiement"),
                                .CommissionTransfert = LireMontant(lecteur, "CommissionTransfert"),
                                .TVA = LireMontant(lecteur, "TVA"),
                                .TTAEnvoi = LireMontant(lecteur, "TTAEnvoi"),
                                .TTAReception = LireMontant(lecteur, "TTAReception"),
                                .TaxeEnvoi = LireMontant(lecteur, "TaxeEnvoi")
                            })
                        End While
                    End Using
                End Using
            End Using

        Catch ex As SqlException
            messageErreur = If(ex.Number = ERREUR_TABLE_ABSENTE,
                               MESSAGE_TABLE_ABSENTE,
                               $"Lecture de l'historique impossible : {ex.Message}")
        Catch ex As InvalidOperationException
            messageErreur = $"Connexion SQL Server indisponible : {ex.Message}"
        End Try

        Return resultat
    End Function

    ''' <summary>
    ''' Première et dernière journées historisées. Servent à proposer d'emblée une période
    ''' pertinente plutôt qu'un intervalle vide.
    ''' </summary>
    ''' <returns>False si l'historique est vide ou illisible.</returns>
    Public Shared Function ObtenirBornes(ByRef premiere As Date, ByRef derniere As Date,
                                         ByRef messageErreur As String) As Boolean

        messageErreur = String.Empty
        premiere = Date.Today
        derniere = Date.Today

        Const requete As String = "SELECT MIN(DateActivite), MAX(DateActivite) FROM " & TABLE_HISTORIQUE

        Try
            Using connexion As SqlConnection = WURepository.CreerConnexion()
                connexion.Open()

                Using commande As New SqlCommand(requete, connexion)
                    Using lecteur As SqlDataReader = commande.ExecuteReader()

                        If Not lecteur.Read() OrElse lecteur.IsDBNull(0) Then
                            Return False ' Historique vide : aucune journée encore comptabilisée.
                        End If

                        premiere = Convert.ToDateTime(lecteur.GetValue(0), Globalization.CultureInfo.InvariantCulture).Date
                        derniere = Convert.ToDateTime(lecteur.GetValue(1), Globalization.CultureInfo.InvariantCulture).Date
                    End Using
                End Using
            End Using

        Catch ex As SqlException
            messageErreur = If(ex.Number = ERREUR_TABLE_ABSENTE,
                               MESSAGE_TABLE_ABSENTE,
                               $"Lecture de l'historique impossible : {ex.Message}")
            Return False
        Catch ex As InvalidOperationException
            messageErreur = $"Connexion SQL Server indisponible : {ex.Message}"
            Return False
        End Try

        Return True
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
        Return Convert.ToInt32(lecteur.GetValue(index), Globalization.CultureInfo.InvariantCulture)
    End Function

    Private Shared Function LireMontant(lecteur As SqlDataReader, colonne As String) As Decimal
        Dim index As Integer = lecteur.GetOrdinal(colonne)
        If lecteur.IsDBNull(index) Then Return 0D
        Return Convert.ToDecimal(lecteur.GetValue(index), Globalization.CultureInfo.InvariantCulture)
    End Function

#End Region

End Class
