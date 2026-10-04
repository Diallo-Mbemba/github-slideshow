Option Strict On
Option Explicit On

Imports System.Data
Imports System.Data.SqlClient

''' <summary>
''' La conservation des écarts de change : la parité en vigueur, le détail par transaction,
''' les lignes écartées et les pièces produites.
'''
''' POURQUOI UN DÉPÔT SÉPARÉ DE ChangeService ET DE PieceChangeService. Ces deux-là sont
''' PURS : ils ne savent pas qu'une base existe. C'est ce qui permet à l'écran de contrôle de
''' les exercer sans rien écrire, et c'est une propriété qu'on perd dès qu'on y glisse un
''' SqlCommand. Tout l'accès à la base est donc ici, et nulle part ailleurs.
'''
''' TOUT S'ÉCRIT EN UNE SEULE TRANSACTION. Le détail, les exclusions et les pièces d'un
''' rapport forment un tout : conserver le détail sans les pièces laisserait une journée
''' calculée mais non comptabilisée, et conserver les pièces sans le détail laisserait une
''' écriture que plus rien ne justifie. L'un sans l'autre est pire que rien du tout.
'''
''' RIEN NE S'ÉCRASE. Les autres dépôts du projet suppriment la journée avant de la réécrire —
''' c'est leur façon de rendre un retraitement possible. Ici, NON : un index unique sur
''' (MTCN, date, sens) refuse la même transaction deux fois, et le refus est le comportement
''' voulu. Un rapport rechargé par erreur doublerait le gain de change de la journée, et les
''' deux pièces seraient équilibrées — personne ne s'en apercevrait. Défaire un enregistrement
''' est un acte d'administration, tracé, pas un effet de bord d'un second clic.
''' </summary>
Public NotInheritable Class ChangeRepository

    Private Sub New()
    End Sub

    Private Const TABLE_PARITE As String = "T_PariteChangeWU"
    Private Const TABLE_ECART As String = "T_EcartChangeWU"
    Private Const TABLE_EXCLUSION As String = "T_ExclusionChangeWU"
    Private Const TABLE_PIECE As String = "T_PieceChangeWU"

    ''' <summary>Code d'erreur SQL Server signalant une table absente (« Invalid object name »).</summary>
    Private Const ERREUR_TABLE_ABSENTE As Integer = 208

    ''' <summary>Violation d'une contrainte unique (2627) ou d'un index unique (2601).</summary>
    Private Const ERREUR_CONTRAINTE_UNIQUE As Integer = 2627
    Private Const ERREUR_INDEX_UNIQUE As Integer = 2601

    ''' <summary>Droit refusé sur un objet (SELECT, INSERT, DELETE).</summary>
    Private Const ERREUR_DROIT_REFUSE As Integer = 229

    Public Const MESSAGE_TABLES_ABSENTES As String =
        "Les tables des écarts de change n'existent pas encore dans la base." & vbCrLf & vbCrLf &
        "Exécutez le script Scripts\00_InstallationComplete.sql : il les crée, amorce la parité et " &
        "accorde les droits." & vbCrLf & vbCrLf &
        "Tant qu'il n'a pas été exécuté, le calcul et la pièce restent consultables à l'écran, " &
        "mais rien ne peut être conservé."

#Region "Parité"

    ''' <summary>
    ''' La parité applicable à une journée : celle dont la date d'effet est la plus récente
    ''' parmi celles qui précèdent cette journée.
    '''
    ''' LA CONSTANTE DU CODE EST LE REPLI, ET ELLE N'EST PAS UN PIS-ALLER. Base injoignable,
    ''' script 21 non exécuté, table vide : la parité fixe du franc CFA s'applique, et c'est la
    ''' bonne réponse — elle n'a pas bougé depuis 1999. Le repli est signalé dans
    ''' messageErreur pour que l'écran puisse le dire, jamais pour arrêter un calcul.
    ''' </summary>
    Public Shared Function PariteEnVigueur(jour As Date, ByRef messageErreur As String) As Decimal

        messageErreur = String.Empty

        Const requete As String =
            "SELECT TOP 1 Parite FROM " & TABLE_PARITE & " " &
            "WHERE DeviseSource = @source AND DeviseCible = @cible AND DateEffet <= @jour " &
            "ORDER BY DateEffet DESC"

        Try
            Using connexion As SqlConnection = WURepository.CreerConnexion()
                connexion.Open()

                Using commande As New SqlCommand(requete, connexion)

                    commande.Parameters.Add("@source", SqlDbType.NVarChar, 3).Value = ConstantesWU.DEVISE_EURO
                    commande.Parameters.Add("@cible", SqlDbType.NVarChar, 3).Value = ConstantesWU.DEVISE_FCFA
                    commande.Parameters.Add("@jour", SqlDbType.Date).Value = jour.Date

                    Dim valeur As Object = commande.ExecuteScalar()

                    If valeur Is Nothing OrElse valeur Is DBNull.Value Then
                        messageErreur = $"Aucune parité {ConstantesWU.DEVISE_EURO}/{ConstantesWU.DEVISE_FCFA} " &
                                        $"en vigueur au {jour:dd/MM/yyyy} dans {TABLE_PARITE} : " &
                                        "la parité fixe du code est employée."
                        Return ConstantesWU.TAUX_CONVERSION
                    End If

                    Dim parite As Decimal = WUReportService.ToDecimalSafe(valeur)

                    ' Une parité nulle ou négative est refusée par une contrainte de la base.
                    ' Le contrôle est répété ici parce qu'une base restaurée d'avant le script
                    ' 21 pourrait ne pas la porter, et qu'une parité à zéro rendrait le gain de
                    ' change égal au chiffre d'affaires.
                    If parite <= 0D Then
                        messageErreur = $"La parité lue dans {TABLE_PARITE} est invalide ({parite}) : " &
                                        "la parité fixe du code est employée."
                        Return ConstantesWU.TAUX_CONVERSION
                    End If

                    Return parite
                End Using
            End Using

        Catch ex As SqlException
            messageErreur = If(ex.Number = ERREUR_TABLE_ABSENTE,
                               MESSAGE_TABLES_ABSENTES,
                               $"Lecture de {TABLE_PARITE} impossible : {ex.Message}") &
                            Environment.NewLine & "La parité fixe du code est employée."
            Return ConstantesWU.TAUX_CONVERSION

        Catch ex As InvalidOperationException
            messageErreur = $"Connexion SQL Server indisponible : {ex.Message}" &
                            Environment.NewLine & "La parité fixe du code est employée."
            Return ConstantesWU.TAUX_CONVERSION
        End Try
    End Function

#End Region

#Region "Contrôle avant écriture"

    ''' <summary>
    ''' Combien des transactions de ce résultat sont DÉJÀ conservées en base.
    '''
    ''' POURQUOI DEMANDER AVANT D'ÉCRIRE, PUISQUE LA BASE REFUSERAIT. Parce que le refus de la
    ''' base arrive au milieu de l'écriture, sur une transaction quelconque, et ne dit pas
    ''' combien d'autres sont concernées : l'agent lirait « violation d'index unique » sur un
    ''' MTCN au hasard. Demander d'abord permet de dire « ce rapport a déjà été enregistré, ses
    ''' 2 367 transactions sont en base » — ou « 3 sur 2 367 », qui est un tout autre problème
    ''' et appelle une tout autre réponse.
    '''
    ''' Retourne -1 si la question n'a pas pu être posée (table absente, base injoignable) :
    ''' l'appelant distingue ainsi « aucune » de « je ne sais pas ».
    ''' </summary>
    Public Shared Function TransactionsDejaConservees(resultat As ResultatChangeWU,
                                                      ByRef messageErreur As String) As Integer

        messageErreur = String.Empty

        If resultat Is Nothing OrElse resultat.Ecarts.Count = 0 Then Return 0

        Const requete As String =
            "SELECT COUNT(*) FROM " & TABLE_ECART & " " &
            "WHERE Mtcn = @mtcn AND DateReglement = @jour AND Sens = @sens"

        Dim deja As Integer = 0

        Try
            Using connexion As SqlConnection = WURepository.CreerConnexion()
                connexion.Open()

                For Each ligne As EcartChangeWU In resultat.Ecarts

                    If Not ligne.DateReglement.HasValue Then Continue For

                    Using commande As New SqlCommand(requete, connexion)
                        commande.Parameters.Add("@mtcn", SqlDbType.NVarChar, 30).Value = ligne.Mtcn
                        commande.Parameters.Add("@jour", SqlDbType.Date).Value = ligne.DateReglement.Value.Date
                        commande.Parameters.Add("@sens", SqlDbType.NChar, 1).Value = ligne.Sens

                        If Convert.ToInt32(commande.ExecuteScalar()) > 0 Then deja += 1
                    End Using
                Next
            End Using

        Catch ex As SqlException
            messageErreur = If(ex.Number = ERREUR_TABLE_ABSENTE,
                               MESSAGE_TABLES_ABSENTES,
                               $"Interrogation de {TABLE_ECART} impossible : {ex.Message}")
            Return -1

        Catch ex As InvalidOperationException
            messageErreur = $"Connexion SQL Server indisponible : {ex.Message}"
            Return -1
        End Try

        Return deja
    End Function

#End Region

#Region "Écriture"

    ''' <summary>
    ''' Conserve un calcul et ses pièces : le détail, les lignes écartées et les écritures,
    ''' en une seule transaction.
    ''' </summary>
    ''' <param name="resultat">Le calcul, tel que ChangeService l'a rendu.</param>
    ''' <param name="pieces">Les pièces, telles que PieceChangeService les a construites.</param>
    ''' <param name="messageErreur">Le motif de l'échec, le cas échéant.</param>
    ''' <returns>True si tout a été écrit. False si rien ne l'a été — jamais un entre-deux.</returns>
    Public Shared Function Conserver(resultat As ResultatChangeWU,
                                     pieces As List(Of PieceChangeWU),
                                     ByRef messageErreur As String) As Boolean

        messageErreur = String.Empty

        If resultat Is Nothing OrElse resultat.Ecarts.Count = 0 Then
            messageErreur = "Aucun écart de change à conserver."
            Return False
        End If

        If pieces Is Nothing OrElse pieces.Count = 0 Then
            messageErreur = "Aucune pièce à conserver : produisez d'abord la pièce de change."
            Return False
        End If

        ' UNE PIÈCE DÉSÉQUILIBRÉE NE S'ÉCRIT PAS. Le contrôle est déjà fait à la construction ;
        ' il est refait ici parce que c'est ici que l'écriture devient définitive, et qu'un
        ' contrôle qui n'est pas au dernier verrou n'est qu'une recommandation.
        For Each piece As PieceChangeWU In pieces
            If Not piece.EstEquilibree Then
                messageErreur = $"La pièce {piece.CleGroupe} n'est pas équilibrée " &
                                $"({piece.TotalDebit} au débit, {piece.TotalCredit} au crédit) : " &
                                "rien n'a été enregistré."
                Return False
            End If
        Next

        Try
            Using connexion As SqlConnection = WURepository.CreerConnexion()
                connexion.Open()

                Using transaction As SqlTransaction = connexion.BeginTransaction()
                    Try
                        EcrireLeDetail(connexion, transaction, resultat)
                        EcrireLesExclusions(connexion, transaction, resultat)
                        EcrireLesPieces(connexion, transaction, pieces)

                        transaction.Commit()

                    Catch
                        ' Le Rollback peut lui-même échouer si la connexion est tombée ; son
                        ' échec ne doit pas masquer l'erreur d'origine, qui est celle qu'on
                        ' remonte.
                        Try
                            transaction.Rollback()
                        Catch
                        End Try

                        Throw
                    End Try
                End Using
            End Using

        Catch ex As SqlException
            messageErreur = MessageDUneErreurSql(ex)
            Return False

        Catch ex As InvalidOperationException
            messageErreur = $"Connexion SQL Server indisponible : {ex.Message}" &
                            Environment.NewLine & "Rien n'a été enregistré."
            Return False
        End Try

        Return True
    End Function

    ''' <summary>
    ''' Traduit une erreur SQL Server en phrase utile. Les trois cas rencontrés en pratique
    ''' méritent chacun leur message : ils appellent trois gestes différents.
    ''' </summary>
    Private Shared Function MessageDUneErreurSql(ex As SqlException) As String

        Select Case ex.Number

            Case ERREUR_TABLE_ABSENTE
                Return MESSAGE_TABLES_ABSENTES

            Case ERREUR_CONTRAINTE_UNIQUE, ERREUR_INDEX_UNIQUE
                Return "Ce rapport — ou une partie de ses transactions — a DÉJÀ été enregistré." &
                       Environment.NewLine & Environment.NewLine &
                       "La base refuse d'enregistrer deux fois la même transaction : sans ce refus, " &
                       "le gain de change de la journée serait doublé et les deux pièces seraient " &
                       "équilibrées — rien ne l'aurait signalé." & Environment.NewLine & Environment.NewLine &
                       "Rien n'a été enregistré. Si un retraitement est réellement nécessaire, " &
                       "l'administrateur doit d'abord défaire l'enregistrement précédent." &
                       Environment.NewLine & Environment.NewLine &
                       $"Détail SQL Server : {ex.Message}"

            Case ERREUR_DROIT_REFUSE
                Return "Droit refusé sur les tables des écarts de change." & Environment.NewLine & Environment.NewLine &
                       "Faites exécuter Scripts\00_InstallationComplete.sql par l'informatique : il accorde " &
                       "les droits en même temps qu'il crée les tables." & Environment.NewLine & Environment.NewLine &
                       $"Détail SQL Server : {ex.Message}"

            Case Else
                Return $"Enregistrement impossible : {ex.Message}" &
                       Environment.NewLine & "Rien n'a été enregistré."
        End Select
    End Function

    Private Shared Sub EcrireLeDetail(connexion As SqlConnection, transaction As SqlTransaction,
                                      resultat As ResultatChangeWU)

        Const insertion As String =
            "INSERT INTO " & TABLE_ECART & " (Mtcn, Account, DateReglement, Sens, CodeProduit, " &
            "Statut, DeviseLocale, MontantLocal, ClearPrincipalLoc, ClearFxLoc, MontantEnDevise, " &
            "Parite, ContreValeur, Ecart, Nature, FichierSource, DateEnregistrement, EnregistrePar) " &
            "VALUES (@mtcn, @account, @jour, @sens, @produit, @statut, @devise, @local, " &
            "@principalDevise, @changeDevise, @devisePayee, @parite, @contreValeur, @ecart, " &
            "@nature, @fichier, GETDATE(), @auteur)"

        For Each ligne As EcartChangeWU In resultat.Ecarts

            ' Une transaction sans date de règlement n'entre dans aucune pièce : la conserver
            ' laisserait en base un écart que rien ne comptabilise. PieceChangeService avertit
            ' déjà l'utilisateur de leur nombre.
            If Not ligne.DateReglement.HasValue Then Continue For

            Using commande As New SqlCommand(insertion, connexion, transaction)

                commande.Parameters.Add("@mtcn", SqlDbType.NVarChar, 30).Value = ligne.Mtcn
                commande.Parameters.Add("@account", SqlDbType.NVarChar, 20).Value = Texte(ligne.Account)
                commande.Parameters.Add("@jour", SqlDbType.Date).Value = ligne.DateReglement.Value.Date
                commande.Parameters.Add("@sens", SqlDbType.NChar, 1).Value = ligne.Sens
                commande.Parameters.Add("@produit", SqlDbType.NVarChar, 10).Value = Texte(ligne.CodeProduit)
                commande.Parameters.Add("@statut", SqlDbType.NVarChar, 10).Value = Texte(ligne.Statut)
                commande.Parameters.Add("@devise", SqlDbType.NVarChar, 10).Value = Texte(ligne.DeviseLocale)
                AjouterDecimal(commande, "@local", ligne.MontantLocal, MONTANT_PRECISION, MONTANT_ECHELLE)
                AjouterDecimal(commande, "@principalDevise", ligne.ClearPrincipalLoc, MONTANT_PRECISION, MONTANT_ECHELLE)
                AjouterDecimal(commande, "@changeDevise", ligne.ClearFxLoc, MONTANT_PRECISION, MONTANT_ECHELLE)
                AjouterDecimal(commande, "@devisePayee", ligne.MontantEnDevise, MONTANT_PRECISION, MONTANT_ECHELLE)
                AjouterDecimal(commande, "@parite", ligne.Parite, PARITE_PRECISION, PARITE_ECHELLE)
                AjouterDecimal(commande, "@contreValeur", ligne.ContreValeur, MONTANT_PRECISION, MONTANT_ECHELLE)
                AjouterDecimal(commande, "@ecart", ligne.Ecart, MONTANT_PRECISION, MONTANT_ECHELLE)
                commande.Parameters.Add("@nature", SqlDbType.NVarChar, 10).Value = ligne.NatureLisible
                commande.Parameters.Add("@fichier", SqlDbType.NVarChar, 260).Value = Texte(resultat.FichierSource)
                commande.Parameters.Add("@auteur", SqlDbType.NVarChar, 50).Value = SessionWU.Auteur

                commande.ExecuteNonQuery()
            End Using
        Next
    End Sub

    Private Shared Sub EcrireLesExclusions(connexion As SqlConnection, transaction As SqlTransaction,
                                           resultat As ResultatChangeWU)

        Const insertion As String =
            "INSERT INTO " & TABLE_EXCLUSION & " (FichierSource, NumeroDeLigne, Mtcn, Motif, " &
            "Detail, DateEnregistrement, EnregistrePar) " &
            "VALUES (@fichier, @numero, @mtcn, @motif, @detail, GETDATE(), @auteur)"

        For Each exclusion As ExclusionChangeWU In resultat.Exclusions

            Using commande As New SqlCommand(insertion, connexion, transaction)

                commande.Parameters.Add("@fichier", SqlDbType.NVarChar, 260).Value = Texte(resultat.FichierSource)
                commande.Parameters.Add("@numero", SqlDbType.Int).Value = exclusion.NumeroDeLigne
                commande.Parameters.Add("@mtcn", SqlDbType.NVarChar, 30).Value = Texte(exclusion.Mtcn)
                commande.Parameters.Add("@motif", SqlDbType.NVarChar, 255).Value = Tronquer(exclusion.Motif, 255)
                commande.Parameters.Add("@detail", SqlDbType.NVarChar, 255).Value = Tronquer(exclusion.Detail, 255)
                commande.Parameters.Add("@auteur", SqlDbType.NVarChar, 50).Value = SessionWU.Auteur

                commande.ExecuteNonQuery()
            End Using
        Next
    End Sub

    Private Shared Sub EcrireLesPieces(connexion As SqlConnection, transaction As SqlTransaction,
                                       pieces As List(Of PieceChangeWU))

        Const insertion As String =
            "INSERT INTO " & TABLE_PIECE & " (CleGroupe, DateReglement, Sens, CodeProduit, " &
            "Account, Mtcn, Ligne, Compte, Libelle, Debit, Credit, CodeAgence, NombreTransactions, " &
            "Parite, FichierSource, DateEnregistrement, EnregistrePar) " &
            "VALUES (@cle, @jour, @sens, @produit, @account, @mtcn, @ligne, @compte, @libelle, " &
            "@debit, @credit, @codeAgence, @transactions, @parite, @fichier, GETDATE(), @auteur)"

        For Each piece As PieceChangeWU In pieces

            If piece.Lignes Is Nothing Then Continue For

            Dim rang As Integer = 0

            For Each ecriture As DataRow In piece.Lignes.Rows

                rang += 1

                Using commande As New SqlCommand(insertion, connexion, transaction)

                    commande.Parameters.Add("@cle", SqlDbType.NVarChar, 60).Value = piece.CleGroupe
                    commande.Parameters.Add("@jour", SqlDbType.Date).Value = piece.DateReglement.Date
                    commande.Parameters.Add("@sens", SqlDbType.NVarChar, 10).Value = Texte(piece.Sens)
                    commande.Parameters.Add("@produit", SqlDbType.NVarChar, 10).Value = Texte(piece.CodeProduit)
                    commande.Parameters.Add("@account", SqlDbType.NVarChar, 20).Value = Texte(piece.Account)
                    commande.Parameters.Add("@mtcn", SqlDbType.NVarChar, 30).Value = Texte(piece.Mtcn)
                    commande.Parameters.Add("@ligne", SqlDbType.Int).Value = rang
                    commande.Parameters.Add("@compte", SqlDbType.NVarChar, 50).Value =
                        Convert.ToString(ecriture("Compte"))
                    commande.Parameters.Add("@libelle", SqlDbType.NVarChar, 255).Value =
                        Tronquer(Convert.ToString(ecriture("Libelle")), 255)
                    commande.Parameters.Add("@debit", SqlDbType.BigInt).Value = Convert.ToInt64(ecriture("Debit"))
                    commande.Parameters.Add("@credit", SqlDbType.BigInt).Value = Convert.ToInt64(ecriture("Credit"))
                    commande.Parameters.Add("@codeAgence", SqlDbType.NVarChar, 10).Value =
                        Convert.ToString(ecriture("CodeAgence"))
                    commande.Parameters.Add("@transactions", SqlDbType.Int).Value = piece.NombreTransactions
                    AjouterDecimal(commande, "@parite", piece.Parite, PARITE_PRECISION, PARITE_ECHELLE)
                    commande.Parameters.Add("@fichier", SqlDbType.NVarChar, 260).Value = Texte(piece.FichierSource)
                    commande.Parameters.Add("@auteur", SqlDbType.NVarChar, 50).Value = SessionWU.Auteur

                    commande.ExecuteNonQuery()
                End Using
            Next
        Next
    End Sub

#End Region

#Region "Lecture des pièces conservées"

    ''' <summary>
    ''' Les pièces de change conservées sur une période, une ligne par écriture.
    '''
    ''' Rendue sous forme de DataTable, et non d'objets : c'est ce que la grille attend, et
    ''' aucun calcul ne s'appuie dessus — seul l'affichage.
    ''' </summary>
    Public Shared Function ListerLesEcritures(debut As Date, fin As Date,
                                              ByRef messageErreur As String) As DataTable

        messageErreur = String.Empty

        Const requete As String =
            "SELECT CleGroupe, DateReglement, Sens, CodeProduit, Account, Mtcn, " &
            "Ligne, Compte, Libelle, " &
            "Debit, Credit, CodeAgence, NombreTransactions, FichierSource, " &
            "DateEnregistrement, EnregistrePar " &
            "FROM " & TABLE_PIECE & " " &
            "WHERE DateReglement >= @debut AND DateReglement <= @fin " &
            "ORDER BY DateReglement, CleGroupe, Ligne"

        Dim table As New DataTable("EcrituresChange")

        Try
            Using connexion As SqlConnection = WURepository.CreerConnexion()
                connexion.Open()

                Using commande As New SqlCommand(requete, connexion)

                    commande.Parameters.Add("@debut", SqlDbType.Date).Value = debut.Date
                    commande.Parameters.Add("@fin", SqlDbType.Date).Value = fin.Date

                    Using lecteur As SqlDataReader = commande.ExecuteReader()
                        table.Load(lecteur)
                    End Using
                End Using
            End Using

        Catch ex As SqlException
            messageErreur = If(ex.Number = ERREUR_TABLE_ABSENTE,
                               MESSAGE_TABLES_ABSENTES,
                               $"Lecture de {TABLE_PIECE} impossible : {ex.Message}")

        Catch ex As InvalidOperationException
            messageErreur = $"Connexion SQL Server indisponible : {ex.Message}"
        End Try

        Return table
    End Function

#End Region

#Region "Utilitaires"

    ''' <summary>
    ''' Pose un paramètre décimal AVEC SA PRÉCISION ET SON ÉCHELLE.
    '''
    ''' CE N'EST PAS UNE FORMALITÉ, ET L'OUBLI NE SE VOIT PAS. Un paramètre SqlDbType.Decimal
    ''' dont l'échelle n'est pas déclarée peut être transmis avec une échelle de zéro : le
    ''' montant 577,39 arrive alors en base comme 577, sans erreur, sans avertissement, et la
    ''' contre-valeur conservée ne correspond plus à l'écart conservé juste à côté. La pièce
    ''' resterait équilibrée — elle est construite à partir des écarts, pas des montants lus —
    ''' et le défaut n'apparaîtrait qu'au premier rapprochement, des mois plus tard.
    '''
    ''' Les valeurs déclarées suivent exactement les colonnes du script 21 : DECIMAL(19,4) pour
    ''' les montants, qui est la précision du rapport Western Union, et DECIMAL(18,6) pour la
    ''' parité.
    ''' </summary>
    Private Shared Sub AjouterDecimal(commande As SqlCommand, nom As String, valeur As Decimal,
                                      precision As Byte, echelle As Byte)

        Dim parametre As SqlParameter = commande.Parameters.Add(nom, SqlDbType.Decimal)
        parametre.Precision = precision
        parametre.Scale = echelle
        parametre.Value = valeur
    End Sub

    ''' <summary>Précision et échelle des colonnes de montant : DECIMAL(19,4).</summary>
    Private Const MONTANT_PRECISION As Byte = 19
    Private Const MONTANT_ECHELLE As Byte = 4

    ''' <summary>Précision et échelle de la parité : DECIMAL(18,6).</summary>
    Private Const PARITE_PRECISION As Byte = 18
    Private Const PARITE_ECHELLE As Byte = 6

    ''' <summary>Une chaîne jamais Nothing, pour un paramètre SQL qui n'accepte pas Nothing.</summary>
    Private Shared Function Texte(valeur As String) As String
        Return If(valeur, String.Empty)
    End Function

    ''' <summary>
    ''' Tronque un texte à la longueur de sa colonne.
    '''
    ''' POURQUOI TRONQUER ICI PLUTÔT QUE LAISSER SQL SERVER REFUSER. Les motifs d'exclusion et
    ''' les libellés sont composés par le code et mesurent cent caractères au plus ; aucun
    ''' n'approche la limite. Mais si l'un d'eux s'allongeait un jour, SQL Server rejetterait
    ''' TOUTE la transaction — le détail, les exclusions et les pièces — pour un motif
    ''' d'exclusion trop bavard. Perdre deux mots d'explication vaut mieux que perdre
    ''' l'enregistrement d'une journée.
    ''' </summary>
    Private Shared Function Tronquer(valeur As String, longueur As Integer) As String

        Dim texte As String = If(valeur, String.Empty)
        If texte.Length <= longueur Then Return texte

        Return texte.Substring(0, longueur)
    End Function

#End Region

End Class
