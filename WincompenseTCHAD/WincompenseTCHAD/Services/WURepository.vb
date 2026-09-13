Option Strict On
Option Explicit On

Imports System.Configuration
Imports System.Data
Imports System.Data.SqlClient

''' <summary>
''' Accès aux paramètres de points de vente (sous-agents / agences propres) dans SQL Server Express.
''' L'Account est toujours utilisé comme identifiant de recherche (jamais codeagence).
''' Aucune exception SQL ne doit remonter jusqu'à l'interface : toute erreur est capturée
''' et reportée sur l'objet CalculWU via ErreurSQL / MessageErreurSQL.
''' </summary>
Public NotInheritable Class WURepository

    Private Sub New()
    End Sub

    ''' <summary>
    ''' Construit la chaîne de connexion SQL Server à utiliser.
    ''' Lue depuis App.config (clé "GWC_WINCOMPENSE_ETD") si disponible, sinon valeur par défaut
    ''' pointant vers .\SQLEXPRESS / GWC_WINCOMPENSE_ETD en authentification Windows intégrée.
    ''' </summary>
    Public Shared Function ObtenirChaineConnexion() As String
        Try
            Dim config As ConnectionStringSettings = ConfigurationManager.ConnectionStrings("GWC_WINCOMPENSE_ETD")
            If config IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(config.ConnectionString) Then
                Return config.ConnectionString
            End If
        Catch
            ' En cas de problème de lecture de configuration, on retombe sur la valeur par défaut ci-dessous.
        End Try

        Return "Server=.\SQLEXPRESS;Database=GWC_WINCOMPENSE_ETD;Integrated Security=True;Connect Timeout=10;"
    End Function

    ''' <summary>
    ''' Ouvre une nouvelle connexion SQL Server. L'appelant est responsable de la fermer (Using).
    ''' </summary>
    Public Shared Function CreerConnexion() As SqlConnection
        Return New SqlConnection(ObtenirChaineConnexion())
    End Function

    ''' <summary>
    ''' Recherche les paramètres d'un Account dans SQL Server, dans l'ordre imposé :
    ''' 1) T_Pdv_SA (sous-agent) via Code_Pdv = Account
    ''' 2) T_Pdv_EC (agence propre) via Codesite = Account
    ''' 3) à défaut, TypePdv = "INCONNU"
    ''' Ne lève jamais d'exception : toute erreur SQL est capturée et posée sur le CalculWU retourné.
    ''' </summary>
    ''' <param name="account">Identifiant Account à rechercher.</param>
    ''' <param name="connexion">Connexion SQL Server déjà ouverte, réutilisée pour toutes les recherches.</param>
    Public Shared Function ChargerParametresAccount(account As String, connexion As SqlConnection) As CalculWU

        Dim resultat As New CalculWU() With {
            .Account = If(account, String.Empty).Trim()
        }

        If String.IsNullOrWhiteSpace(resultat.Account) Then
            resultat.TypePdv = "INCONNU"
            resultat.Designation = "ACCOUNT NON RENSEIGNE"
            resultat.DonneesManquantes = True
            Return resultat
        End If

        Try
            If RechercherSousAgent(resultat, connexion) Then
                Return resultat
            End If

            If RechercherAgencePropre(resultat, connexion) Then
                Return resultat
            End If

            ' Aucune correspondance trouvée dans T_Pdv_SA ni T_Pdv_EC.
            resultat.TypePdv = "INCONNU"
            resultat.Designation = "ACCOUNT NON PARAMETRE"

        Catch ex As SqlException
            resultat.ErreurSQL = True
            resultat.MessageErreurSQL = $"Erreur SQL lors de la recherche de l'Account {resultat.Account} : {ex.Message}"
            resultat.TypePdv = "INCONNU"
            resultat.Designation = "ERREUR SQL - PARAMETRES NON RECUPERES"
        Catch ex As InvalidOperationException
            ' Ex : connexion fermée / base inaccessible.
            resultat.ErreurSQL = True
            resultat.MessageErreurSQL = $"Connexion SQL Server indisponible pour l'Account {resultat.Account} : {ex.Message}"
            resultat.TypePdv = "INCONNU"
            resultat.Designation = "BASE SQL INACCESSIBLE"
        End Try

        Return resultat
    End Function

    ''' <summary>Recherche l'Account comme sous-agent dans T_Pdv_SA. Retourne True si trouvé.</summary>
    Private Shared Function RechercherSousAgent(calc As CalculWU, connexion As SqlConnection) As Boolean

        Const requete As String =
            "SELECT Code_Pdv, Designationagence, GroupeStatistique, Taux, " &
            "CompteCompense, CompteCommission, codeagence " &
            "FROM T_Pdv_SA WHERE Code_Pdv = @Account"

        Using commande As New SqlCommand(requete, connexion)
            commande.Parameters.Add("@Account", SqlDbType.VarChar, 50).Value = calc.Account

            Using lecteur As SqlDataReader = commande.ExecuteReader()
                If lecteur.Read() Then
                    calc.TypePdv = "SA"
                    calc.Designation = LireChaine(lecteur, "Designationagence")
                    calc.GroupeStatistique = LireChaine(lecteur, "GroupeStatistique")
                    calc.CodeAgence = LireChaine(lecteur, "codeagence")
                    calc.CompteCompense = LireChaine(lecteur, "CompteCompense")
                    calc.CompteCommission = LireChaine(lecteur, "CompteCommission")
                    calc.TauxSA = LireDecimal(lecteur, "Taux")

                    If String.IsNullOrWhiteSpace(calc.CompteCompense) OrElse String.IsNullOrWhiteSpace(calc.CompteCommission) Then
                        calc.DonneesManquantes = True
                    End If

                    Return True
                End If
            End Using
        End Using

        Return False
    End Function

    ''' <summary>Recherche l'Account comme agence propre dans T_Pdv_EC. Retourne True si trouvé.</summary>
    Private Shared Function RechercherAgencePropre(calc As CalculWU, connexion As SqlConnection) As Boolean

        Const requete As String =
            "SELECT Codesite, Designationagence, [CodeAgenc-Voyager] " &
            "FROM T_Pdv_EC WHERE Codesite = @Account"

        Using commande As New SqlCommand(requete, connexion)
            commande.Parameters.Add("@Account", SqlDbType.VarChar, 50).Value = calc.Account

            Using lecteur As SqlDataReader = commande.ExecuteReader()
                If lecteur.Read() Then
                    calc.TypePdv = "EC"
                    calc.Designation = LireChaine(lecteur, "Designationagence")
                    calc.CodeAgence = LireChaine(lecteur, "CodeAgenc-Voyager")
                    ' Agence propre : la banque conserve 100% des commissions (pas de compte de compensation SA).
                    calc.TauxSA = 0D
                    Return True
                End If
            End Using
        End Using

        Return False
    End Function

#Region "Comptes systèmes de la pièce comptable (table SystemeWU)"

    ''' <summary>Nom de la table SQL Server portant le paramétrage des comptes comptables.</summary>
    Private Const TABLE_SYSTEME As String = "SystemeWU"

    ''' <summary>
    ''' Colonnes de SystemeWU exploitées par l'application. Les autres colonnes de la table
    ''' (Passif, Actif, Cpte_Charge_Publicitaire, Cpte_Gainde_Change, Cpte_Envoi_agence,
    ''' Cpte_Paiement_agence) ne sont ni lues ni écrites : elles ne concernent pas la pièce
    ''' comptable Western Union et un enregistrement ne doit jamais les altérer.
    ''' </summary>
    Private Const COLONNES_SYSTEME As String =
        "Cpte_PositionNette, Cpte_attenteDEBIT, Cpte_attenteCREDIT, " &
        "Cpte_Produit, Cpte_Produit_Envoi, Cpte_Produit_Paiement, " &
        "Tthu, Tob, Cpte_Envoi, Cpte_Paiement, code"

    ''' <summary>
    ''' Charge les comptes comptables depuis la table SystemeWU.
    '''
    ''' Ne lève jamais d'exception : si la base est inaccessible, si la table est absente ou
    ''' vide, les comptes PAR DÉFAUT du code sont retournés (ChargeDepuisBase = False) et la
    ''' raison est posée dans messageErreur. L'application reste ainsi utilisable, avec
    ''' exactement le comportement qu'elle avait avant le paramétrage.
    '''
    ''' Une colonne vide ou NULL dans la base ne remplace pas la valeur par défaut : mieux vaut
    ''' un compte par défaut connu qu'une écriture comptable sans numéro de compte. Le
    ''' formulaire de paramétrage permet ensuite de compléter la ligne.
    ''' </summary>
    ''' <param name="messageErreur">Vide si les comptes viennent bien de la base.</param>
    Public Shared Function ChargerComptesSysteme(ByRef messageErreur As String) As ComptesSystemeWU

        messageErreur = String.Empty

        ' Valeurs par défaut : conservées telles quelles pour toute colonne non exploitable.
        Dim comptes As New ComptesSystemeWU()

        Try
            Using connexion As SqlConnection = CreerConnexion()
                connexion.Open()

                Dim requete As String = $"SELECT TOP 1 {COLONNES_SYSTEME} FROM {TABLE_SYSTEME} ORDER BY code"

                Using commande As New SqlCommand(requete, connexion)
                    Using lecteur As SqlDataReader = commande.ExecuteReader()

                        If Not lecteur.Read() Then
                            messageErreur = $"La table {TABLE_SYSTEME} ne contient aucune ligne de paramétrage : " &
                                            "les comptes par défaut de l'application sont utilisés."
                            Return comptes
                        End If

                        AffecterSiRenseigne(lecteur, "Cpte_PositionNette", Sub(v) comptes.CompteCourant = v)

                        ' Le compte inter bancaire est porté par deux colonnes que la ligne de
                        ' paramétrage tient identiques : la colonne DÉBIT fait foi, la colonne
                        ' CRÉDIT ne prend le relais que si la première n'est pas renseignée.
                        Dim interBancaire As String = LireColonneSysteme(lecteur, "Cpte_attenteDEBIT")
                        If interBancaire.Length = 0 Then
                            interBancaire = LireColonneSysteme(lecteur, "Cpte_attenteCREDIT")
                        End If
                        If interBancaire.Length > 0 Then
                            comptes.CompteInterBancaire = interBancaire
                        End If

                        AffecterSiRenseigne(lecteur, "Cpte_Produit", Sub(v) comptes.CommissionTransfertBanque = v)
                        AffecterSiRenseigne(lecteur, "Cpte_Produit_Envoi", Sub(v) comptes.CommissionEnvoiBanque = v)
                        AffecterSiRenseigne(lecteur, "Cpte_Produit_Paiement", Sub(v) comptes.CommissionPaiementBanque = v)
                        AffecterSiRenseigne(lecteur, "Tthu", Sub(v) comptes.ImpotsTaxeEnvoi = v)
                        AffecterSiRenseigne(lecteur, "Tob", Sub(v) comptes.TVACollectee = v)
                        AffecterSiRenseigne(lecteur, "Cpte_Envoi", Sub(v) comptes.TTAEnvoi = v)
                        AffecterSiRenseigne(lecteur, "Cpte_Paiement", Sub(v) comptes.TTAReception = v)
                        AffecterSiRenseigne(lecteur, "code", Sub(v) comptes.CodeParametrage = v)

                        comptes.ChargeDepuisBase = True
                    End Using
                End Using
            End Using

        Catch ex As SqlException
            messageErreur = $"Lecture de la table {TABLE_SYSTEME} impossible : {ex.Message}" &
                            Environment.NewLine & "Les comptes par défaut de l'application sont utilisés."
        Catch ex As InvalidOperationException
            messageErreur = $"Connexion SQL Server indisponible : {ex.Message}" &
                            Environment.NewLine & "Les comptes par défaut de l'application sont utilisés."
        End Try

        Return comptes
    End Function

    ''' <summary>
    ''' Enregistre les comptes comptables dans la table SystemeWU.
    '''
    ''' Seules les colonnes exploitées par l'application sont mises à jour : les autres colonnes
    ''' de la ligne restent intactes. Le compte inter bancaire est écrit dans les DEUX colonnes
    ''' Cpte_attenteDEBIT et Cpte_attenteCREDIT, que le formulaire présente comme un seul champ
    ''' et que la ligne de paramétrage existante tient identiques.
    '''
    ''' La ligne visée est celle dont la colonne « code » vaut CodeParametrage. Si aucun code
    ''' n'a pu être lu, la mise à jour n'est acceptée que si la table ne contient qu'une seule
    ''' ligne — jamais question de modifier une ligne au hasard dans une table partagée.
    ''' </summary>
    ''' <returns>True si l'enregistrement a bien eu lieu.</returns>
    Public Shared Function EnregistrerComptesSysteme(comptes As ComptesSystemeWU, ByRef messageErreur As String) As Boolean

        messageErreur = String.Empty

        If comptes Is Nothing Then
            messageErreur = "Aucun compte à enregistrer."
            Return False
        End If

        Try
            Using connexion As SqlConnection = CreerConnexion()
                connexion.Open()

                Dim filtre As String = String.Empty
                Dim codeVise As String = If(comptes.CodeParametrage, String.Empty).Trim()

                If codeVise.Length > 0 Then
                    ' La colonne code est de type nchar : comparaison sur la valeur ajustée par SQL Server.
                    filtre = " WHERE RTRIM(code) = @code"
                Else
                    Dim nombreLignes As Integer = CompterLignesSysteme(connexion)

                    If nombreLignes = 0 Then
                        messageErreur = $"La table {TABLE_SYSTEME} est vide : aucune ligne de paramétrage à mettre à jour." &
                                        Environment.NewLine & Environment.NewLine &
                                        "Créez-la au préalable en exécutant le script Scripts\03_SystemeWU.sql."
                        Return False
                    End If

                    If nombreLignes > 1 Then
                        messageErreur = $"La table {TABLE_SYSTEME} contient {nombreLignes} lignes, mais celle qui a été lue " &
                                        "ne porte aucun code." & Environment.NewLine & Environment.NewLine &
                                        "Impossible de déterminer sans ambiguïté la ligne à mettre à jour : " &
                                        "renseignez la colonne « code » de la ligne de paramétrage Western Union."
                        Return False
                    End If
                End If

                Dim requete As String =
                    $"UPDATE {TABLE_SYSTEME} SET " &
                    "Cpte_PositionNette = @compteCourant, " &
                    "Cpte_attenteDEBIT = @interBancaire, " &
                    "Cpte_attenteCREDIT = @interBancaire, " &
                    "Cpte_Produit = @commissionTransfert, " &
                    "Cpte_Produit_Envoi = @commissionEnvoi, " &
                    "Cpte_Produit_Paiement = @commissionPaiement, " &
                    "Tthu = @impotsTaxeEnvoi, " &
                    "Tob = @tva, " &
                    "Cpte_Envoi = @ttaEnvoi, " &
                    "Cpte_Paiement = @ttaReception, " &
                    "DateModification = GETDATE(), ModifiePar = @auteur" & filtre

                Using commande As New SqlCommand(requete, connexion)

                    commande.Parameters.Add("@compteCourant", SqlDbType.NVarChar, 255).Value = comptes.CompteCourant
                    commande.Parameters.Add("@interBancaire", SqlDbType.NVarChar, 255).Value = comptes.CompteInterBancaire
                    commande.Parameters.Add("@commissionTransfert", SqlDbType.NVarChar, 255).Value = comptes.CommissionTransfertBanque
                    commande.Parameters.Add("@commissionEnvoi", SqlDbType.NVarChar, 255).Value = comptes.CommissionEnvoiBanque
                    commande.Parameters.Add("@commissionPaiement", SqlDbType.NVarChar, 255).Value = comptes.CommissionPaiementBanque
                    commande.Parameters.Add("@impotsTaxeEnvoi", SqlDbType.NVarChar, 255).Value = comptes.ImpotsTaxeEnvoi
                    commande.Parameters.Add("@tva", SqlDbType.NVarChar, 255).Value = comptes.TVACollectee
                    commande.Parameters.Add("@ttaEnvoi", SqlDbType.NVarChar, 255).Value = comptes.TTAEnvoi
                    commande.Parameters.Add("@ttaReception", SqlDbType.NVarChar, 255).Value = comptes.TTAReception
                    commande.Parameters.Add("@auteur", SqlDbType.NVarChar, 50).Value = SessionWU.Auteur

                    If codeVise.Length > 0 Then
                        commande.Parameters.Add("@code", SqlDbType.NVarChar, 10).Value = codeVise
                    End If

                    Dim lignesModifiees As Integer = commande.ExecuteNonQuery()

                    If lignesModifiees = 0 Then
                        messageErreur = $"Aucune ligne mise à jour dans {TABLE_SYSTEME}" &
                                        If(codeVise.Length > 0, $" pour le code « {codeVise} »", String.Empty) & "." &
                                        Environment.NewLine &
                                        "La ligne de paramétrage a peut-être été supprimée ou son code modifié."
                        Return False
                    End If
                End Using
            End Using

        Catch ex As SqlException
            messageErreur = $"Enregistrement impossible dans {TABLE_SYSTEME} : {ex.Message}"
            Return False
        Catch ex As InvalidOperationException
            messageErreur = $"Connexion SQL Server indisponible : {ex.Message}"
            Return False
        End Try

        Return True
    End Function

    ''' <summary>Nombre de lignes de la table de paramétrage.</summary>
    Private Shared Function CompterLignesSysteme(connexion As SqlConnection) As Integer
        Using commande As New SqlCommand($"SELECT COUNT(*) FROM {TABLE_SYSTEME}", connexion)
            Dim valeur As Object = commande.ExecuteScalar()
            If valeur Is Nothing OrElse valeur Is DBNull.Value Then Return 0
            Return Convert.ToInt32(valeur, Globalization.CultureInfo.InvariantCulture)
        End Using
    End Function

    ''' <summary>
    ''' Affecte la valeur d'une colonne UNIQUEMENT si elle est renseignée dans la base.
    ''' Une colonne absente de la table, NULL ou vide laisse donc intacte la valeur par défaut.
    ''' </summary>
    Private Shared Sub AffecterSiRenseigne(lecteur As SqlDataReader, colonne As String, affectation As Action(Of String))
        Dim valeur As String = LireColonneSysteme(lecteur, colonne)
        If valeur.Length > 0 Then
            affectation(valeur)
        End If
    End Sub

    ''' <summary>
    ''' Lit une colonne de SystemeWU, ou une chaîne vide si elle est absente de la table, NULL
    ''' ou vide. Une colonne absente n'est pas une erreur : la table est partagée et peut ne pas
    ''' comporter exactement les mêmes colonnes d'un environnement à l'autre.
    ''' </summary>
    Private Shared Function LireColonneSysteme(lecteur As SqlDataReader, colonne As String) As String

        Dim index As Integer
        Try
            index = lecteur.GetOrdinal(colonne)
        Catch ex As IndexOutOfRangeException
            Return String.Empty
        End Try

        If lecteur.IsDBNull(index) Then Return String.Empty

        Return lecteur.GetValue(index).ToString().Trim()
    End Function

#End Region

#Region "Utilitaires de lecture robustes (gestion DBNull)"

    Private Shared Function LireChaine(lecteur As SqlDataReader, colonne As String) As String
        Dim index As Integer = lecteur.GetOrdinal(colonne)
        If lecteur.IsDBNull(index) Then Return String.Empty
        Return lecteur.GetValue(index).ToString().Trim()
    End Function

    Private Shared Function LireDecimal(lecteur As SqlDataReader, colonne As String) As Decimal
        Dim index As Integer = lecteur.GetOrdinal(colonne)
        If lecteur.IsDBNull(index) Then Return 0D
        Return WUReportService.ToDecimalSafe(lecteur.GetValue(index))
    End Function

#End Region

End Class
