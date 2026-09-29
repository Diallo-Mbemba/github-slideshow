Option Strict On
Option Explicit On

Imports System.Data
Imports System.Data.SqlClient
Imports System.Drawing

''' <summary>
''' Lecture et enregistrement des produits de transfert (table T_ProduitTransfert).
'''
''' NE LÈVE JAMAIS D'EXCEPTION À LA LECTURE. Table absente, base injoignable, table vide : la
''' liste écrite en dur — Western Union et Ria — est retournée, et la raison est posée dans
''' messageErreur. L'application démarre et compense sans cette table, exactement comme avant
''' qu'elle n'existe. C'est la règle de la maison, déjà appliquée aux comptes systèmes et au
''' barème des taxes.
'''
''' CE QUE CE DÉPÔT N'ÉCRIT PAS
'''
''' L'état « traité par l'application » n'a aucune colonne ici, et aucune requête ne le touche.
''' C'est le code qui le sait, et lui seul : voir ProduitTransfert.EstTraiteParLApplication.
''' Un administrateur ne peut donc pas déclarer un produit disponible et faire lire à son écran
''' de traitement les rapports et les comptes de Western Union.
'''
''' AUCUNE SUPPRESSION. Le dépôt n'expose pas de DELETE, et le script ne l'accorde à personne.
''' Un produit supprimé laisserait son historique, ses pièces et son référentiel orphelins.
''' Mettre hors service le retire de la fenêtre de choix sans rien effacer.
''' </summary>
Public NotInheritable Class ProduitRepository

    Private Sub New()
    End Sub

    Private Const TABLE As String = "T_ProduitTransfert"

    ''' <summary>Code d'erreur SQL Server signalant une table absente (« Invalid object name »).</summary>
    Private Const ERREUR_TABLE_ABSENTE As Integer = 208

    ''' <summary>Code d'erreur SQL Server signalant une violation de clé (doublon).</summary>
    Private Const ERREUR_DOUBLON As Integer = 2627

    Public Const MESSAGE_TABLE_ABSENTE As String =
        "La table T_ProduitTransfert n'existe pas encore dans la base." & vbCrLf & vbCrLf &
        "Exécutez le script Scripts\25_ProduitsTransfert.sql : il la crée et l'amorce avec " &
        "Western Union et Ria." & vbCrLf &
        "Tant qu'elle est absente, l'application emploie sa liste interne — les deux mêmes " &
        "produits — et l'écran d'administration ne peut rien enregistrer."

#Region "Lecture"

    ''' <summary>
    ''' Charge les produits depuis la table et les installe dans ProduitTransfert.
    '''
    ''' Retourne Vrai si la liste vient bien de la base. Faux si la liste de secours a pris le
    ''' relais — la raison est alors dans messageErreur.
    ''' </summary>
    Public Shared Function Charger(ByRef messageErreur As String) As Boolean

        messageErreur = String.Empty

        Dim lus As New List(Of ProduitTransfert)

        Try
            Using connexion As SqlConnection = WURepository.CreerConnexion()
                connexion.Open()

                Const requete As String =
                    "SELECT Code, Nom, Description, CouleurRouge, CouleurVert, CouleurBleu, " &
                    "Ordre, EnService, CreePar, DateCreation, ModifiePar, DateModification " &
                    "FROM " & TABLE & " ORDER BY Ordre, Nom"

                Using commande As New SqlCommand(requete, connexion)
                    Using lecteur As SqlDataReader = commande.ExecuteReader()
                        While lecteur.Read()

                            Dim produit As ProduitTransfert = LireUnProduit(lecteur)
                            If produit IsNot Nothing Then lus.Add(produit)
                        End While
                    End Using
                End Using
            End Using

        Catch ex As SqlException

            messageErreur = If(ex.Number = ERREUR_TABLE_ABSENTE,
                               MESSAGE_TABLE_ABSENTE,
                               $"Lecture de la table {TABLE} impossible : {ex.Message}" &
                               Environment.NewLine &
                               "La liste interne des produits est employée.")

            ProduitTransfert.Charger(ProduitTransfert.ListeDeSecours())
            Return False

        Catch ex As Exception

            messageErreur = $"Lecture des produits impossible : {ex.Message}" &
                            Environment.NewLine & "La liste interne des produits est employée."

            ProduitTransfert.Charger(ProduitTransfert.ListeDeSecours())
            Return False
        End Try

        ' Une table présente mais VIDE n'est pas une liste vide : c'est un amorçage qui n'a pas
        ' eu lieu. Sans ce garde-fou, la fenêtre de choix s'ouvrirait sans aucune ligne et
        ' personne ne pourrait plus travailler.
        If lus.Count = 0 Then
            messageErreur = $"La table {TABLE} est vide. La liste interne des produits est employée." &
                            Environment.NewLine &
                            "Rejouez Scripts\25_ProduitsTransfert.sql pour l'amorcer."

            ProduitTransfert.Charger(ProduitTransfert.ListeDeSecours())
            Return False
        End If

        ProduitTransfert.Charger(lus)
        Return True
    End Function

    ''' <summary>Construit un produit depuis la ligne courante, ou Nothing si elle est inutilisable.</summary>
    Private Shared Function LireUnProduit(lecteur As SqlDataReader) As ProduitTransfert

        Dim code As String = TexteDe(lecteur, 0)
        If code.Length = 0 Then Return Nothing

        Dim produit As New ProduitTransfert(code)

        produit.Nom = TexteDe(lecteur, 1)
        produit.Description = TexteDe(lecteur, 2)

        ' Une composante manquante ne fait pas une couleur noire : la couleur par défaut du
        ' produit est conservée. Les trois sont donc lues ensemble ou pas du tout.
        If Not lecteur.IsDBNull(3) AndAlso Not lecteur.IsDBNull(4) AndAlso Not lecteur.IsDBNull(5) Then
            produit.Couleur = Color.FromArgb(Convert.ToInt32(lecteur.GetValue(3)),
                                             Convert.ToInt32(lecteur.GetValue(4)),
                                             Convert.ToInt32(lecteur.GetValue(5)))
        End If

        If Not lecteur.IsDBNull(6) Then produit.Ordre = Convert.ToInt32(lecteur.GetValue(6))
        If Not lecteur.IsDBNull(7) Then produit.EnService = Convert.ToBoolean(lecteur.GetValue(7))

        produit.CreePar = TexteDe(lecteur, 8)
        If Not lecteur.IsDBNull(9) Then produit.DateCreation = Convert.ToDateTime(lecteur.GetValue(9))

        produit.ModifiePar = TexteDe(lecteur, 10)
        If Not lecteur.IsDBNull(11) Then produit.DateModification = Convert.ToDateTime(lecteur.GetValue(11))

        Return produit
    End Function

    Private Shared Function TexteDe(lecteur As SqlDataReader, colonne As Integer) As String

        If lecteur.IsDBNull(colonne) Then Return String.Empty
        Return Convert.ToString(lecteur.GetValue(colonne)).Trim()
    End Function

#End Region

#Region "Écriture"

    ''' <summary>
    ''' Crée un produit. Retourne Vrai si la ligne est posée ; messageErreur dit pourquoi
    ''' sinon — un code déjà pris, par exemple.
    ''' </summary>
    Public Shared Function Creer(produit As ProduitTransfert, ByRef messageErreur As String) As Boolean

        messageErreur = String.Empty
        If produit Is Nothing Then
            messageErreur = "Aucun produit à créer."
            Return False
        End If

        Try
            Using connexion As SqlConnection = WURepository.CreerConnexion()
                connexion.Open()

                Const requete As String =
                    "INSERT INTO " & TABLE &
                    " (Code, Nom, Description, CouleurRouge, CouleurVert, CouleurBleu, " &
                    "Ordre, EnService, CreePar, DateCreation) VALUES " &
                    "(@code, @nom, @description, @rouge, @vert, @bleu, @ordre, @enService, " &
                    "@auteur, GETDATE())"

                Using commande As New SqlCommand(requete, connexion)

                    PoserLesParametres(commande, produit)
                    commande.Parameters.AddWithValue("@code", produit.Code)
                    commande.Parameters.AddWithValue("@auteur", SessionWU.Identifiant)

                    commande.ExecuteNonQuery()
                End Using
            End Using

            Return True

        Catch ex As SqlException

            messageErreur = MessageDeLErreur(ex, $"Le produit « {produit.Code} » existe déjà.")
            Return False

        Catch ex As Exception

            messageErreur = $"Création du produit impossible : {ex.Message}"
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Enregistre les modifications d'un produit. LE CODE N'EST PAS MODIFIABLE : il sert de
    ''' clé à la mise à jour, il nomme le fichier du logo et il nommera les tables du produit.
    ''' </summary>
    Public Shared Function Enregistrer(produit As ProduitTransfert, ByRef messageErreur As String) As Boolean

        messageErreur = String.Empty
        If produit Is Nothing Then
            messageErreur = "Aucun produit à enregistrer."
            Return False
        End If

        Try
            Using connexion As SqlConnection = WURepository.CreerConnexion()
                connexion.Open()

                Const requete As String =
                    "UPDATE " & TABLE & " SET Nom = @nom, Description = @description, " &
                    "CouleurRouge = @rouge, CouleurVert = @vert, CouleurBleu = @bleu, " &
                    "Ordre = @ordre, EnService = @enService, " &
                    "ModifiePar = @auteur, DateModification = GETDATE() WHERE Code = @code"

                Using commande As New SqlCommand(requete, connexion)

                    PoserLesParametres(commande, produit)
                    commande.Parameters.AddWithValue("@code", produit.Code)
                    commande.Parameters.AddWithValue("@auteur", SessionWU.Identifiant)

                    ' Zéro ligne touchée : la ligne n'existe pas encore. Cela arrive sur une
                    ' base dont l'amorçage n'a pas été joué, et le dire vaut mieux que
                    ' d'annoncer un enregistrement qui n'a rien enregistré.
                    If commande.ExecuteNonQuery() = 0 Then
                        messageErreur = $"Le produit « {produit.Code} » n'existe pas dans la table."
                        Return False
                    End If
                End Using
            End Using

            Return True

        Catch ex As Exception

            messageErreur = $"Enregistrement du produit impossible : {ex.Message}"
            Return False
        End Try
    End Function

    ''' <summary>Les paramètres communs à la création et à la mise à jour.</summary>
    Private Shared Sub PoserLesParametres(commande As SqlCommand, produit As ProduitTransfert)

        commande.Parameters.AddWithValue("@nom", produit.Nom)
        commande.Parameters.AddWithValue("@description",
                                         If(String.IsNullOrEmpty(produit.Description),
                                            CObj(DBNull.Value), CObj(produit.Description)))

        commande.Parameters.AddWithValue("@rouge", CInt(produit.Couleur.R))
        commande.Parameters.AddWithValue("@vert", CInt(produit.Couleur.G))
        commande.Parameters.AddWithValue("@bleu", CInt(produit.Couleur.B))

        commande.Parameters.AddWithValue("@ordre", produit.Ordre)
        commande.Parameters.AddWithValue("@enService", produit.EnService)
    End Sub

    ''' <summary>Le message qui convient à une erreur SQL : le doublon a le sien.</summary>
    Private Shared Function MessageDeLErreur(erreur As SqlException, messageDuDoublon As String) As String

        If erreur.Number = ERREUR_DOUBLON Then Return messageDuDoublon
        If erreur.Number = ERREUR_TABLE_ABSENTE Then Return MESSAGE_TABLE_ABSENTE

        Return $"Écriture dans la table {TABLE} impossible : {erreur.Message}"
    End Function

#End Region

End Class
