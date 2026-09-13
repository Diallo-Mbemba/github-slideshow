Option Strict On
Option Explicit On

Imports System.Data
Imports System.Data.SqlClient

''' <summary>
''' File des écritures proposées sur le référentiel des points de vente (T_DemandeWU).
'''
''' Un INPUTER y dépose ; un AUTHORIZER y décide. Ce n'est qu'à l'autorisation que l'écriture
''' est réellement portée dans T_Pdv_SA, T_Pdv_EC ou T_GroupeStatistique — et l'application et
''' la décision tiennent dans une seule transaction : jamais l'une sans l'autre.
'''
''' Les écritures directes restent dans PdvRepository, qui devient la couche d'application :
''' plus aucun écran ne l'appelle en écriture, seule l'autorisation y conduit.
''' </summary>
Public NotInheritable Class DemandeRepository

    Private Sub New()
    End Sub

    Private Const TABLE As String = "T_DemandeWU"

    Private Const ERREUR_TABLE_ABSENTE As Integer = 208
    Private Const ERREUR_CLE_DUPLIQUEE As Integer = 2627
    Private Const ERREUR_INDEX_UNIQUE As Integer = 2601
    Private Const ERREUR_CONTRAINTE As Integer = 547

    Public Const MESSAGE_TABLE_ABSENTE As String =
        "La table T_DemandeWU n'existe pas encore dans la base." & vbCrLf & vbCrLf &
        "Exécutez le script Scripts\09_Demandes.sql : il crée la file des demandes et ajoute " &
        "la colonne Fonction à la table des utilisateurs."

    Private Const COLONNES As String =
        "IdDemande, TypeObjet, Operation, Statut, Cle, Designation, GroupeStatistique, " &
        "CompteActivite, CompteCommission, Taux, CodeRattachement, " &
        "SaisiPar, DateSaisie, DecidePar, DateDecision, MotifRejet"

#Region "Dépôt d'une demande"

    ''' <summary>
    ''' Dépose une demande. Elle n'a aucun effet tant qu'un authorizer ne l'a pas autorisée.
    ''' </summary>
    Public Shared Function Soumettre(demande As DemandeWU, ByRef messageErreur As String) As Boolean

        messageErreur = String.Empty

        If demande Is Nothing Then
            messageErreur = "Aucune demande à déposer."
            Return False
        End If

        If Not SessionWU.PeutSaisirLesPointsDeVente Then
            messageErreur = "Seul un utilisateur ayant la fonction « inputer » peut saisir " &
                            "une modification du référentiel."
            Return False
        End If

        If String.IsNullOrWhiteSpace(demande.Cle) Then
            messageErreur = "La demande ne désigne aucun objet."
            Return False
        End If

        Const requete As String =
            "INSERT INTO " & TABLE & " (TypeObjet, Operation, Statut, Cle, Designation, " &
            "GroupeStatistique, CompteActivite, CompteCommission, Taux, CodeRattachement, " &
            "SaisiPar, DateSaisie) " &
            "VALUES (@objet, @operation, @statut, @cle, @designation, @groupe, @activite, " &
            "@commission, @taux, @rattachement, @saisiPar, GETDATE())"

        Try
            Using connexion As SqlConnection = WURepository.CreerConnexion()
                connexion.Open()

                Using commande As New SqlCommand(requete, connexion)
                    AjouterParametres(commande, demande)
                    commande.ExecuteNonQuery()
                End Using
            End Using

        Catch ex As SqlException
            messageErreur = MessageErreur(ex, demande)
            Return False
        Catch ex As InvalidOperationException
            messageErreur = $"Connexion SQL Server indisponible : {ex.Message}"
            Return False
        End Try

        Return True
    End Function

    Private Shared Sub AjouterParametres(commande As SqlCommand, demande As DemandeWU)

        commande.Parameters.Add("@objet", SqlDbType.NVarChar, 20).Value = DemandeWU.LibelleDepuisObjet(demande.TypeObjet)
        commande.Parameters.Add("@operation", SqlDbType.NVarChar, 20).Value = DemandeWU.LibelleDepuisOperation(demande.Operation)
        commande.Parameters.Add("@statut", SqlDbType.NVarChar, 20).Value = DemandeWU.STATUT_EN_ATTENTE
        commande.Parameters.Add("@cle", SqlDbType.NVarChar, 255).Value = demande.Cle.Trim()
        commande.Parameters.Add("@designation", SqlDbType.NVarChar, 255).Value = If(demande.Designation, String.Empty)
        commande.Parameters.Add("@groupe", SqlDbType.NVarChar, 255).Value = If(demande.GroupeStatistique, String.Empty)
        commande.Parameters.Add("@activite", SqlDbType.NVarChar, 255).Value = If(demande.CompteActivite, String.Empty)
        commande.Parameters.Add("@commission", SqlDbType.NVarChar, 255).Value = If(demande.CompteCommission, String.Empty)

        ' La colonne Taux est de type DECIMAL(4,2), comme dans les tables cibles : précision et
        ' échelle cadrées explicitement, valeur arrondie côté application.
        Dim parametreTaux As SqlParameter = commande.Parameters.Add("@taux", SqlDbType.Decimal)
        parametreTaux.Precision = 4
        parametreTaux.Scale = 2
        parametreTaux.Value = Decimal.Round(demande.Taux, 2)

        commande.Parameters.Add("@rattachement", SqlDbType.NVarChar, 255).Value = If(demande.CodeRattachement, String.Empty)
        commande.Parameters.Add("@saisiPar", SqlDbType.NVarChar, 50).Value = SessionWU.Auteur
    End Sub

    Private Shared Function MessageErreur(ex As SqlException, demande As DemandeWU) As String

        If ex.Number = ERREUR_TABLE_ABSENTE Then Return MESSAGE_TABLE_ABSENTE

        If ex.Number = ERREUR_CLE_DUPLIQUEE OrElse ex.Number = ERREUR_INDEX_UNIQUE Then
            Return $"Une demande est déjà en attente sur {demande.LibelleObjet.ToLowerInvariant()} " &
                   $"« {demande.Cle} »." & Environment.NewLine & Environment.NewLine &
                   "Elle doit être autorisée ou rejetée avant qu'une autre puisse être déposée : " &
                   "deux demandes contradictoires sur le même objet s'appliqueraient sinon dans " &
                   "l'ordre où l'authorizer les traite."
        End If

        Return $"Dépôt de la demande impossible : {ex.Message}"
    End Function

#End Region

#Region "Lecture"

    ''' <summary>Nombre de demandes en attente. Sert au compteur du menu.</summary>
    Public Shared Function CompterEnAttente() As Integer

        Const requete As String = "SELECT COUNT(*) FROM " & TABLE & " WHERE Statut = @statut"

        Try
            Using connexion As SqlConnection = WURepository.CreerConnexion()
                connexion.Open()

                Using commande As New SqlCommand(requete, connexion)
                    commande.Parameters.Add("@statut", SqlDbType.NVarChar, 20).Value = DemandeWU.STATUT_EN_ATTENTE

                    Dim valeur As Object = commande.ExecuteScalar()
                    If valeur Is Nothing OrElse valeur Is DBNull.Value Then Return 0
                    Return Convert.ToInt32(valeur, Globalization.CultureInfo.InvariantCulture)
                End Using
            End Using

        Catch ex As SqlException
            ' Table absente ou base injoignable : le compteur se tait plutôt que d'alarmer.
            Return 0
        Catch ex As InvalidOperationException
            Return 0
        End Try
    End Function

    ''' <summary>Demandes en attente, la plus ancienne d'abord.</summary>
    Public Shared Function ListerEnAttente(ByRef messageErreur As String) As List(Of DemandeWU)
        Return Lister(DemandeWU.STATUT_EN_ATTENTE, 0, messageErreur)
    End Function

    ''' <summary>Demandes déjà décidées, les plus récentes d'abord.</summary>
    Public Shared Function ListerHistorique(nombreMaximum As Integer, ByRef messageErreur As String) As List(Of DemandeWU)
        Return Lister(String.Empty, nombreMaximum, messageErreur)
    End Function

    ''' <param name="statut">Statut recherché, ou chaîne vide pour les demandes déjà décidées.</param>
    ''' <param name="nombreMaximum">0 pour ne pas limiter.</param>
    Private Shared Function Lister(statut As String, nombreMaximum As Integer,
                                   ByRef messageErreur As String) As List(Of DemandeWU)

        messageErreur = String.Empty
        Dim resultat As New List(Of DemandeWU)

        Dim tete As String = If(nombreMaximum > 0, $"TOP {nombreMaximum} ", String.Empty)

        Dim requete As String =
            If(statut.Length > 0,
               $"SELECT {tete}{COLONNES} FROM {TABLE} WHERE Statut = @statut ORDER BY DateSaisie",
               $"SELECT {tete}{COLONNES} FROM {TABLE} WHERE Statut <> @statut ORDER BY DateDecision DESC")

        Try
            Using connexion As SqlConnection = WURepository.CreerConnexion()
                connexion.Open()

                Using commande As New SqlCommand(requete, connexion)
                    commande.Parameters.Add("@statut", SqlDbType.NVarChar, 20).Value =
                        If(statut.Length > 0, statut, DemandeWU.STATUT_EN_ATTENTE)

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
                               $"Lecture des demandes impossible : {ex.Message}")
        Catch ex As InvalidOperationException
            messageErreur = $"Connexion SQL Server indisponible : {ex.Message}"
        End Try

        Return resultat
    End Function

    Private Shared Function Construire(lecteur As SqlDataReader) As DemandeWU

        Return New DemandeWU() With {
            .IdDemande = LireEntierLong(lecteur, "IdDemande"),
            .TypeObjet = DemandeWU.ObjetDepuisLibelle(LireChaine(lecteur, "TypeObjet")),
            .Operation = DemandeWU.OperationDepuisLibelle(LireChaine(lecteur, "Operation")),
            .Statut = DemandeWU.StatutDepuisLibelle(LireChaine(lecteur, "Statut")),
            .Cle = LireChaine(lecteur, "Cle"),
            .Designation = LireChaine(lecteur, "Designation"),
            .GroupeStatistique = LireChaine(lecteur, "GroupeStatistique"),
            .CompteActivite = LireChaine(lecteur, "CompteActivite"),
            .CompteCommission = LireChaine(lecteur, "CompteCommission"),
            .Taux = LireDecimal(lecteur, "Taux"),
            .CodeRattachement = LireChaine(lecteur, "CodeRattachement"),
            .SaisiPar = LireChaine(lecteur, "SaisiPar"),
            .DateSaisie = LireDate(lecteur, "DateSaisie"),
            .DecidePar = LireChaine(lecteur, "DecidePar"),
            .DateDecision = LireDate(lecteur, "DateDecision"),
            .MotifRejet = LireChaine(lecteur, "MotifRejet")
        }
    End Function

#End Region

#Region "Décision"

    ''' <summary>
    ''' Autorise une demande et porte l'écriture dans la table cible.
    '''
    ''' Les deux se font dans une seule transaction : une demande marquée autorisée alors que
    ''' l'écriture a échoué laisserait croire que la donnée est en base, et personne ne
    ''' reviendrait dessus.
    ''' </summary>
    Public Shared Function Autoriser(demande As DemandeWU, ByRef messageErreur As String) As Boolean

        messageErreur = String.Empty

        If Not ControlerDecision(demande, messageErreur) Then Return False

        Try
            Using connexion As SqlConnection = WURepository.CreerConnexion()
                connexion.Open()

                Using transaction As SqlTransaction = connexion.BeginTransaction()
                    Try
                        ' La demande est verrouillée et relue dans la transaction : sans cela,
                        ' deux authorizers pourraient autoriser la même demande en même temps et
                        ' appliquer deux fois l'écriture.
                        If Not VerrouillerEnAttente(demande.IdDemande, connexion, transaction) Then
                            transaction.Rollback()
                            messageErreur = "Cette demande vient d'être traitée par quelqu'un d'autre." &
                                            Environment.NewLine & "Actualisez la liste."
                            Return False
                        End If

                        If Not Appliquer(demande, connexion, transaction, messageErreur) Then
                            transaction.Rollback()
                            Return False
                        End If

                        Decider(demande.IdDemande, DemandeWU.STATUT_AUTORISE, String.Empty,
                                connexion, transaction)

                        transaction.Commit()

                    Catch
                        transaction.Rollback()
                        Throw
                    End Try
                End Using
            End Using

        Catch ex As SqlException
            messageErreur = If(ex.Number = ERREUR_TABLE_ABSENTE,
                               MESSAGE_TABLE_ABSENTE,
                               $"Autorisation impossible : {ex.Message}")
            Return False
        Catch ex As InvalidOperationException
            messageErreur = $"Connexion SQL Server indisponible : {ex.Message}"
            Return False
        End Try

        Return True
    End Function

    ''' <summary>Rejette une demande. Rien n'est écrit dans les tables cibles.</summary>
    Public Shared Function Rejeter(demande As DemandeWU, motif As String,
                                   ByRef messageErreur As String) As Boolean

        messageErreur = String.Empty

        If Not ControlerDecision(demande, messageErreur) Then Return False

        If String.IsNullOrWhiteSpace(motif) Then
            messageErreur = "Le motif du rejet est obligatoire : sans lui, celui qui a saisi " &
                            "la demande ne sait pas quoi corriger."
            Return False
        End If

        Try
            Using connexion As SqlConnection = WURepository.CreerConnexion()
                connexion.Open()

                Using transaction As SqlTransaction = connexion.BeginTransaction()
                    Try
                        If Not VerrouillerEnAttente(demande.IdDemande, connexion, transaction) Then
                            transaction.Rollback()
                            messageErreur = "Cette demande vient d'être traitée par quelqu'un d'autre." &
                                            Environment.NewLine & "Actualisez la liste."
                            Return False
                        End If

                        Decider(demande.IdDemande, DemandeWU.STATUT_REJETE, motif.Trim(),
                                connexion, transaction)

                        transaction.Commit()

                    Catch
                        transaction.Rollback()
                        Throw
                    End Try
                End Using
            End Using

        Catch ex As SqlException
            messageErreur = If(ex.Number = ERREUR_TABLE_ABSENTE,
                               MESSAGE_TABLE_ABSENTE,
                               $"Rejet impossible : {ex.Message}")
            Return False
        Catch ex As InvalidOperationException
            messageErreur = $"Connexion SQL Server indisponible : {ex.Message}"
            Return False
        End Try

        Return True
    End Function

    ''' <summary>
    ''' Contrôles communs à l'autorisation et au rejet.
    '''
    ''' La règle « personne ne décide de sa propre saisie » est vérifiée ici, et redoublée par
    ''' une contrainte CHECK dans la base : un UPDATE fait à la main dans Management Studio ne
    ''' peut pas davantage la contourner.
    ''' </summary>
    Private Shared Function ControlerDecision(demande As DemandeWU, ByRef messageErreur As String) As Boolean

        If demande Is Nothing Then
            messageErreur = "Aucune demande sélectionnée."
            Return False
        End If

        If Not SessionWU.PeutAutoriserLesPointsDeVente Then
            messageErreur = "Seul un utilisateur ayant la fonction « authorizer » peut décider " &
                            "d'une demande."
            Return False
        End If

        If demande.Statut <> StatutDemandeWU.EnAttente Then
            messageErreur = "Cette demande a déjà été décidée."
            Return False
        End If

        If String.Equals(demande.SaisiPar, SessionWU.Auteur, StringComparison.OrdinalIgnoreCase) Then
            messageErreur = "Vous avez saisi cette demande : vous ne pouvez pas la décider " &
                            "vous-même." & Environment.NewLine & Environment.NewLine &
                            "C'est tout l'objet du double regard — un seul agent ne doit pas " &
                            "pouvoir modifier seul une donnée qui détermine des écritures " &
                            "comptables."
            Return False
        End If

        Return True
    End Function

    ''' <summary>
    ''' Verrouille la demande et vérifie qu'elle est bien encore en attente.
    ''' Retourne faux si elle a été décidée entre-temps.
    ''' </summary>
    Private Shared Function VerrouillerEnAttente(idDemande As Long, connexion As SqlConnection,
                                                 transaction As SqlTransaction) As Boolean

        Const requete As String =
            "SELECT COUNT(*) FROM " & TABLE & " WITH (UPDLOCK, ROWLOCK) " &
            "WHERE IdDemande = @id AND Statut = @statut"

        Using commande As New SqlCommand(requete, connexion, transaction)
            commande.Parameters.Add("@id", SqlDbType.BigInt).Value = idDemande
            commande.Parameters.Add("@statut", SqlDbType.NVarChar, 20).Value = DemandeWU.STATUT_EN_ATTENTE

            Dim valeur As Object = commande.ExecuteScalar()
            If valeur Is Nothing OrElse valeur Is DBNull.Value Then Return False
            Return Convert.ToInt32(valeur, Globalization.CultureInfo.InvariantCulture) > 0
        End Using
    End Function

    Private Shared Sub Decider(idDemande As Long, statut As String, motif As String,
                               connexion As SqlConnection, transaction As SqlTransaction)

        Const requete As String =
            "UPDATE " & TABLE & " SET Statut = @statut, DecidePar = @decidePar, " &
            "DateDecision = GETDATE(), MotifRejet = @motif WHERE IdDemande = @id"

        Using commande As New SqlCommand(requete, connexion, transaction)
            commande.Parameters.Add("@id", SqlDbType.BigInt).Value = idDemande
            commande.Parameters.Add("@statut", SqlDbType.NVarChar, 20).Value = statut
            commande.Parameters.Add("@decidePar", SqlDbType.NVarChar, 50).Value = SessionWU.Auteur
            commande.Parameters.Add("@motif", SqlDbType.NVarChar, 500).Value = If(motif, String.Empty)

            commande.ExecuteNonQuery()
        End Using
    End Sub

#End Region

#Region "Application de l'écriture"

    ''' <summary>
    ''' Porte l'écriture demandée dans sa table cible, dans la transaction de la décision.
    ''' </summary>
    Private Shared Function Appliquer(demande As DemandeWU, connexion As SqlConnection,
                                      transaction As SqlTransaction,
                                      ByRef messageErreur As String) As Boolean

        Select Case demande.TypeObjet

            Case TypeObjetWU.SousAgent
                Return PdvRepository.AppliquerSousAgent(demande, connexion, transaction, messageErreur)

            Case TypeObjetWU.Agence
                Return PdvRepository.AppliquerAgence(demande, connexion, transaction, messageErreur)

            Case TypeObjetWU.Groupe
                Return PdvRepository.AppliquerGroupe(demande, connexion, transaction, messageErreur)

            Case Else
                messageErreur = "Type d'objet inconnu : la demande ne peut pas être appliquée."
                Return False
        End Select
    End Function

#End Region

#Region "Utilitaires internes"

    Private Shared Function LireChaine(lecteur As SqlDataReader, colonne As String) As String
        Dim index As Integer = lecteur.GetOrdinal(colonne)
        If lecteur.IsDBNull(index) Then Return String.Empty
        Return lecteur.GetValue(index).ToString().Trim()
    End Function

    Private Shared Function LireEntierLong(lecteur As SqlDataReader, colonne As String) As Long
        Dim index As Integer = lecteur.GetOrdinal(colonne)
        If lecteur.IsDBNull(index) Then Return 0L
        Return Convert.ToInt64(lecteur.GetValue(index), Globalization.CultureInfo.InvariantCulture)
    End Function

    Private Shared Function LireDecimal(lecteur As SqlDataReader, colonne As String) As Decimal
        Dim index As Integer = lecteur.GetOrdinal(colonne)
        If lecteur.IsDBNull(index) Then Return 0D
        Return Convert.ToDecimal(lecteur.GetValue(index), Globalization.CultureInfo.InvariantCulture)
    End Function

    Private Shared Function LireDate(lecteur As SqlDataReader, colonne As String) As Date?
        Dim index As Integer = lecteur.GetOrdinal(colonne)
        If lecteur.IsDBNull(index) Then Return Nothing
        Return Convert.ToDateTime(lecteur.GetValue(index), Globalization.CultureInfo.InvariantCulture)
    End Function

#End Region

End Class
