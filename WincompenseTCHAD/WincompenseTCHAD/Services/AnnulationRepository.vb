Option Strict On
Option Explicit On

Imports System.Data
Imports System.Data.SqlClient
Imports System.Globalization

''' <summary>
''' Annulation d'une journée comptabilisée : dépôt de la demande, puis déplacement effectif
''' des lignes vers les tables d'archive au moment de l'autorisation.
'''
''' ANNULER N'EST PAS SUPPRIMER. Les lignes de T_HistoriqueWU, T_HistoriqueMTCN et T_PieceWU
''' sont DÉPLACÉES vers T_HistoriqueAnnuleWU, T_HistoriqueMTCNAnnuleWU et T_PieceAnnuleeWU,
''' sous un identifiant d'annulation qui porte le motif et les deux signatures. La journée
''' sort des rapports d'activité — qui lisent les tables vivantes — sans sortir de la mémoire
''' de l'application.
'''
''' POURQUOI LE DÉPLACEMENT EST ENTIÈREMENT EN SQL
'''
''' INSERT ... SELECT puis DELETE : les lignes ne remontent jamais dans l'application. Une
''' journée chargée peut porter plusieurs milliers de MTCN, et les faire voyager pour les
''' réécrire aussitôt n'apporterait rien — sinon le risque qu'une conversion de type en
''' altère une au passage.
'''
''' TOUT TIENT DANS LA TRANSACTION DE L'AUTORISATION
'''
''' L'en-tête, les trois déplacements et la décision sur la demande sont une seule
''' transaction. Une journée à moitié archivée serait comptée deux fois par les rapports,
''' ou pas du tout.
''' </summary>
Public NotInheritable Class AnnulationRepository

    Private Sub New()
    End Sub

    Private Const TABLE_ANNULATION As String = "T_AnnulationWU"
    Private Const TABLE_DEMANDE As String = "T_DemandeWU"

    Private Const HISTORIQUE As String = "T_HistoriqueWU"
    Private Const HISTORIQUE_ARCHIVE As String = "T_HistoriqueAnnuleWU"
    Private Const MTCN As String = "T_HistoriqueMTCN"
    Private Const MTCN_ARCHIVE As String = "T_HistoriqueMTCNAnnuleWU"
    Private Const PIECE As String = "T_PieceWU"
    Private Const PIECE_ARCHIVE As String = "T_PieceAnnuleeWU"

    ''' <summary>Code d'erreur SQL Server signalant une table absente (« Invalid object name »).</summary>
    Private Const ERREUR_TABLE_ABSENTE As Integer = 208

    Public Const MESSAGE_TABLES_ABSENTES As String =
        "Les tables d'annulation n'existent pas encore dans la base." & vbCrLf & vbCrLf &
        "Exécutez le script Scripts\14_AnnulationComptabilisation.sql : il les crée et " &
        "ouvre la file des demandes aux annulations." & vbCrLf &
        "Tant qu'elles sont absentes, une comptabilisation ne peut pas être retirée."

#Region "Ce que contient une journée"

    ''' <summary>
    ''' Ce qu'une annulation retirerait : de quoi le montrer à l'agent AVANT qu'il décide,
    ''' et de quoi le conserver dans l'archive une fois la journée partie.
    ''' </summary>
    Public NotInheritable Class ContenuJournee

        Public Property NombrePdv As Integer = 0
        Public Property NombreTransactions As Integer = 0
        Public Property NombreEcritures As Integer = 0
        Public Property TotalDebit As Long = 0L
        Public Property TotalCredit As Long = 0L

        ''' <summary>
        ''' Vrai si la journée n'a plus rien de comptabilisé. Annuler alors ne retirerait
        ''' rien et laisserait une archive vide, qui ferait croire à un traitement.
        ''' </summary>
        Public ReadOnly Property EstVide As Boolean
            Get
                Return NombrePdv = 0 AndAlso NombreTransactions = 0 AndAlso NombreEcritures = 0
            End Get
        End Property

        Public ReadOnly Property Equilibree As Boolean
            Get
                Return TotalDebit = TotalCredit
            End Get
        End Property

        ''' <summary>Phrase résumant ce qui sera retiré.</summary>
        Public ReadOnly Property Intitule As String
            Get
                Return $"{NombrePdv} point(s) de vente, {NombreTransactions} transaction(s), " &
                       $"{NombreEcritures} écriture(s) — {TotalDebit:N0} FCFA au débit."
            End Get
        End Property
    End Class

    ''' <summary>
    ''' Décrit ce que contient une journée comptabilisée, sur une connexion à part.
    ''' </summary>
    Public Shared Function Decrire(jour As Date, ByRef messageErreur As String) As ContenuJournee

        messageErreur = String.Empty

        Try
            Using connexion As SqlConnection = WURepository.CreerConnexion()
                connexion.Open()
                Return Decrire(jour, connexion, Nothing)
            End Using

        Catch ex As SqlException
            messageErreur = If(ex.Number = ERREUR_TABLE_ABSENTE,
                               HistoriqueRepository.MESSAGE_TABLE_ABSENTE,
                               $"Lecture de la journée du {jour:dd/MM/yyyy} impossible : {ex.Message}")
            Return New ContenuJournee()

        Catch ex As InvalidOperationException
            messageErreur = $"Connexion SQL Server indisponible : {ex.Message}"
            Return New ContenuJournee()
        End Try
    End Function

    ''' <summary>
    ''' Décrit la journée DANS la connexion et la transaction de l'appelant : au moment
    ''' d'autoriser, les chiffres écrits dans l'archive doivent être ceux que l'on déplace,
    ''' et non ceux lus une seconde plus tôt sur une autre connexion.
    ''' </summary>
    Private Shared Function Decrire(jour As Date, connexion As SqlConnection,
                                    transaction As SqlTransaction) As ContenuJournee

        Dim contenu As New ContenuJournee()

        contenu.NombrePdv = Compter(
            "SELECT COUNT(*) FROM " & HISTORIQUE & " WHERE DateActivite = @jour",
            jour, connexion, transaction)

        contenu.NombreTransactions = Compter(
            "SELECT COUNT(*) FROM " & MTCN & " WHERE DateActivite = @jour",
            jour, connexion, transaction)

        Const totaux As String =
            "SELECT ecritures = COUNT(*), " &
            "       totalDebit = ISNULL(SUM(Debit), 0), " &
            "       totalCredit = ISNULL(SUM(Credit), 0) " &
            "FROM   " & PIECE & " WHERE DateActivite = @jour"

        Using commande As New SqlCommand(totaux, connexion, transaction)
            commande.Parameters.Add("@jour", SqlDbType.Date).Value = jour.Date

            Using lecteur As SqlDataReader = commande.ExecuteReader()
                If lecteur.Read() Then
                    contenu.NombreEcritures = Convert.ToInt32(lecteur.GetValue(0), CultureInfo.InvariantCulture)
                    contenu.TotalDebit = Convert.ToInt64(lecteur.GetValue(1), CultureInfo.InvariantCulture)
                    contenu.TotalCredit = Convert.ToInt64(lecteur.GetValue(2), CultureInfo.InvariantCulture)
                End If
            End Using
        End Using

        Return contenu
    End Function

    Private Shared Function Compter(requete As String, jour As Date, connexion As SqlConnection,
                                    transaction As SqlTransaction) As Integer

        Using commande As New SqlCommand(requete, connexion, transaction)
            commande.Parameters.Add("@jour", SqlDbType.Date).Value = jour.Date

            Dim valeur As Object = commande.ExecuteScalar()
            If valeur Is Nothing OrElse valeur Is DBNull.Value Then Return 0
            Return Convert.ToInt32(valeur, CultureInfo.InvariantCulture)
        End Using
    End Function

#End Region

#Region "Dépôt de la demande"

    ''' <summary>
    ''' Dépose une demande d'annulation. Elle n'a AUCUN effet tant qu'un authorizer ne l'a
    ''' pas autorisée : la journée reste dans les rapports jusque-là.
    '''
    ''' Le dépôt emprunte la file du référentiel (T_DemandeWU) : l'authorizer garde un seul
    ''' écran, la règle « personne ne décide de sa propre saisie » s'applique telle quelle,
    ''' et l'index unique sur (TypeObjet, Cle) interdit deux demandes simultanées sur la
    ''' même journée.
    ''' </summary>
    Public Shared Function Proposer(annulation As AnnulationWU, ByRef messageErreur As String) As Boolean

        messageErreur = String.Empty

        If annulation Is Nothing Then
            messageErreur = "Aucune annulation à déposer."
            Return False
        End If

        If annulation.Motif = MotifAnnulationWU.Inconnu Then
            messageErreur = "Le motif de l'annulation est obligatoire." & Environment.NewLine &
                            "Sans lui, l'archive ne dira pas pourquoi la journée a été retirée."
            Return False
        End If

        ' « Autre » n'explique rien par lui-même : il n'a de sens qu'accompagné.
        If annulation.Motif = MotifAnnulationWU.Autre AndAlso
           String.IsNullOrWhiteSpace(annulation.Commentaire) Then

            messageErreur = "Le motif « Autre » demande une explication écrite."
            Return False
        End If

        Dim demande As DemandeWU = DemandeWU.DepuisAnnulation(annulation)
        Return DemandeRepository.Soumettre(demande, messageErreur)
    End Function

    ''' <summary>
    ''' La demande d'annulation en attente sur une journée, ou Nothing.
    '''
    ''' Sert à avertir avant de recomptabiliser : si quelqu'un a demandé le retrait de cette
    ''' journée et qu'on la refait entre-temps, l'authorizer annulerait la NOUVELLE version
    ''' en croyant annuler l'ancienne.
    ''' </summary>
    Public Shared Function DemandeEnAttente(jour As Date) As DemandeWU

        Const lecture As String =
            "SELECT TOP 1 IdDemande, Designation, Commentaire, SaisiPar, DateSaisie " &
            "FROM   " & TABLE_DEMANDE & " " &
            "WHERE  TypeObjet = @objet AND Cle = @cle AND Statut = @statut " &
            "ORDER BY DateSaisie DESC"

        Try
            Using connexion As SqlConnection = WURepository.CreerConnexion()
                connexion.Open()

                Using commande As New SqlCommand(lecture, connexion)

                    commande.Parameters.Add("@objet", SqlDbType.NVarChar, 20).Value = DemandeWU.OBJET_COMPTABILISATION
                    commande.Parameters.Add("@cle", SqlDbType.NVarChar, 255).Value = CleDepuisJour(jour)
                    commande.Parameters.Add("@statut", SqlDbType.NVarChar, 20).Value = DemandeWU.STATUT_EN_ATTENTE

                    Using lecteur As SqlDataReader = commande.ExecuteReader()

                        If Not lecteur.Read() Then Return Nothing

                        Return New DemandeWU() With {
                            .IdDemande = Convert.ToInt64(lecteur.GetValue(0), CultureInfo.InvariantCulture),
                            .TypeObjet = TypeObjetWU.Comptabilisation,
                            .Operation = OperationWU.Annulation,
                            .Statut = StatutDemandeWU.EnAttente,
                            .Cle = CleDepuisJour(jour),
                            .Designation = LireChaine(lecteur, 1),
                            .Commentaire = LireChaine(lecteur, 2),
                            .SaisiPar = LireChaine(lecteur, 3),
                            .DateSaisie = lecteur.GetDateTime(4)
                        }
                    End Using
                End Using
            End Using

        Catch ex As SqlException
            ' Tables absentes ou base injoignable : pas d'avertissement plutôt qu'un faux.
            Return Nothing
        Catch ex As InvalidOperationException
            Return Nothing
        End Try
    End Function

#End Region

#Region "Application de l'annulation"

    ''' <summary>
    ''' Retire effectivement la journée, DANS LA TRANSACTION DE LA DÉCISION.
    '''
    ''' Appelée par DemandeRepository au moment où l'authorizer autorise : l'en-tête
    ''' d'annulation, les trois déplacements et la décision sur la demande forment une seule
    ''' transaction, et aucune des quatre n'existe sans les autres.
    ''' </summary>
    Public Shared Function Appliquer(demande As DemandeWU, connexion As SqlConnection,
                                     transaction As SqlTransaction,
                                     ByRef messageErreur As String) As Boolean

        messageErreur = String.Empty

        Dim jour As Date
        If Not JourDepuisCle(demande.Cle, jour) Then
            messageErreur = $"La demande ne désigne pas une journée exploitable (« {demande.Cle} »)."
            Return False
        End If

        ' Les chiffres écrits dans l'archive sont ceux que l'on déplace, lus dans la même
        ' transaction : entre le dépôt de la demande et son autorisation, la journée a pu
        ' être recomptabilisée.
        Dim contenu As ContenuJournee = Decrire(jour, connexion, transaction)

        If contenu.EstVide Then
            messageErreur =
                $"La journée du {jour:dd/MM/yyyy} n'est plus comptabilisée : il n'y a rien à retirer." &
                Environment.NewLine & Environment.NewLine &
                "Elle a peut-être déjà été annulée depuis le dépôt de cette demande." &
                Environment.NewLine &
                "Rejetez la demande : autoriser créerait une archive vide, qui ferait croire " &
                "à un traitement."
            Return False
        End If

        Dim idAnnulation As Long = EcrireEnTete(demande, jour, contenu, connexion, transaction)

        DeplacerHistorique(idAnnulation, jour, connexion, transaction)
        DeplacerTransactions(idAnnulation, jour, connexion, transaction)
        DeplacerPiece(idAnnulation, jour, connexion, transaction)

        Return True
    End Function

    ''' <summary>Écrit l'en-tête de l'annulation et retourne son identifiant.</summary>
    Private Shared Function EcrireEnTete(demande As DemandeWU, jour As Date, contenu As ContenuJournee,
                                         connexion As SqlConnection,
                                         transaction As SqlTransaction) As Long

        Const insertion As String =
            "INSERT INTO " & TABLE_ANNULATION & " (DateActivite, Motif, Commentaire, " &
            "CoreBankingInjecte, NombrePdv, NombreTransactions, NombreEcritures, " &
            "TotalDebit, TotalCredit, IdDemande, DemandeePar, DateDemande, " &
            "AutoriseePar, DateAutorisation) " &
            "VALUES (@jour, @motif, @commentaire, @coreBanking, @pdv, @transactions, " &
            "@ecritures, @debit, @credit, @idDemande, @demandeePar, @dateDemande, " &
            "@autoriseePar, GETDATE()); " &
            "SELECT CAST(SCOPE_IDENTITY() AS BIGINT);"

        Using commande As New SqlCommand(insertion, connexion, transaction)

            commande.Parameters.Add("@jour", SqlDbType.Date).Value = jour.Date
            commande.Parameters.Add("@motif", SqlDbType.NVarChar, 30).Value =
                If(demande.Designation, String.Empty).Trim().ToUpperInvariant()
            commande.Parameters.Add("@commentaire", SqlDbType.NVarChar, 500).Value =
                If(demande.Commentaire, String.Empty)
            commande.Parameters.Add("@coreBanking", SqlDbType.Bit).Value = demande.CoreBankingInjecte

            commande.Parameters.Add("@pdv", SqlDbType.Int).Value = contenu.NombrePdv
            commande.Parameters.Add("@transactions", SqlDbType.Int).Value = contenu.NombreTransactions
            commande.Parameters.Add("@ecritures", SqlDbType.Int).Value = contenu.NombreEcritures
            commande.Parameters.Add("@debit", SqlDbType.BigInt).Value = contenu.TotalDebit
            commande.Parameters.Add("@credit", SqlDbType.BigInt).Value = contenu.TotalCredit

            commande.Parameters.Add("@idDemande", SqlDbType.BigInt).Value = demande.IdDemande
            commande.Parameters.Add("@demandeePar", SqlDbType.NVarChar, 50).Value =
                If(demande.SaisiPar, String.Empty)
            commande.Parameters.Add("@dateDemande", SqlDbType.DateTime).Value =
                If(demande.DateSaisie.HasValue, CType(demande.DateSaisie.Value, Object), CType(Date.Now, Object))
            commande.Parameters.Add("@autoriseePar", SqlDbType.NVarChar, 50).Value = SessionWU.Auteur

            Dim valeur As Object = commande.ExecuteScalar()
            Return Convert.ToInt64(valeur, CultureInfo.InvariantCulture)
        End Using
    End Function

    Private Shared Sub DeplacerHistorique(idAnnulation As Long, jour As Date,
                                          connexion As SqlConnection, transaction As SqlTransaction)

        Const colonnes As String =
            "DateActivite, Account, Designation, GroupeStatistique, TypePdv, " &
            "NombreEnvois, NombrePaiements, NombreAnnulations, " &
            "PrincipalEnvoi, ChargeEnvoi, Taxes, PrincipalPaye, " &
            "CommissionEnvoi, CommissionPaiement, CommissionTransfert, " &
            "TVA, TTAEnvoi, TTAReception, TaxeEnvoi, " &
            "DateEnregistrement, ComptabilisePar, DateComptabilisation"

        Deplacer(HISTORIQUE, HISTORIQUE_ARCHIVE, colonnes, idAnnulation, jour, connexion, transaction)
    End Sub

    Private Shared Sub DeplacerTransactions(idAnnulation As Long, jour As Date,
                                            connexion As SqlConnection, transaction As SqlTransaction)

        Const colonnes As String =
            "DateActivite, Account, MTCN, Sens, Statut, Montant, " &
            "Designation, GroupeStatistique, TypePdv, " &
            "DateEnregistrement, ComptabilisePar, DateComptabilisation"

        Deplacer(MTCN, MTCN_ARCHIVE, colonnes, idAnnulation, jour, connexion, transaction)
    End Sub

    Private Shared Sub DeplacerPiece(idAnnulation As Long, jour As Date,
                                     connexion As SqlConnection, transaction As SqlTransaction)

        Const colonnes As String =
            "DateActivite, Ligne, Compte, Libelle, Debit, Credit, CodeAgence, " &
            "DateEnregistrement, EnregistrePar"

        Deplacer(PIECE, PIECE_ARCHIVE, colonnes, idAnnulation, jour, connexion, transaction)
    End Sub

    ''' <summary>
    ''' Copie puis efface les lignes d'une journée. Les deux ordres se suivent dans la même
    ''' transaction : une copie sans effacement compterait la journée deux fois.
    '''
    ''' Les noms de tables et de colonnes sont des constantes du programme, jamais une saisie :
    ''' rien de ce que l'utilisateur écrit n'entre dans ces requêtes, seule la journée y entre,
    ''' et elle y entre en paramètre.
    ''' </summary>
    Private Shared Sub Deplacer(source As String, archive As String, colonnes As String,
                                idAnnulation As Long, jour As Date,
                                connexion As SqlConnection, transaction As SqlTransaction)

        Dim copie As String =
            $"INSERT INTO {archive} (IdAnnulation, {colonnes}) " &
            $"SELECT @id, {colonnes} FROM {source} WHERE DateActivite = @jour"

        Using commande As New SqlCommand(copie, connexion, transaction)
            commande.Parameters.Add("@id", SqlDbType.BigInt).Value = idAnnulation
            commande.Parameters.Add("@jour", SqlDbType.Date).Value = jour.Date
            commande.ExecuteNonQuery()
        End Using

        Dim effacement As String = $"DELETE FROM {source} WHERE DateActivite = @jour"

        Using commande As New SqlCommand(effacement, connexion, transaction)
            commande.Parameters.Add("@jour", SqlDbType.Date).Value = jour.Date
            commande.ExecuteNonQuery()
        End Using
    End Sub

#End Region

#Region "La journée, écrite dans la demande"

    ''' <summary>
    ''' La journée telle qu'elle est enregistrée dans la colonne Cle de la file des demandes.
    ''' Format ISO, et non le format français : il se trie, et il ne dépend d'aucune culture.
    ''' </summary>
    Public Shared Function CleDepuisJour(jour As Date) As String
        Return jour.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
    End Function

    ''' <summary>Relit la journée écrite dans la demande. Faux si la clé n'est pas une date.</summary>
    Public Shared Function JourDepuisCle(cle As String, ByRef jour As Date) As Boolean

        jour = Date.MinValue

        Return Date.TryParseExact(If(cle, String.Empty).Trim(), "yyyy-MM-dd",
                                  CultureInfo.InvariantCulture, DateTimeStyles.None, jour)
    End Function

#End Region

#Region "Utilitaires de lecture"

    Private Shared Function LireChaine(lecteur As SqlDataReader, position As Integer) As String
        If lecteur.IsDBNull(position) Then Return String.Empty
        Return lecteur.GetValue(position).ToString().Trim()
    End Function

#End Region

End Class
