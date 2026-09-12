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
