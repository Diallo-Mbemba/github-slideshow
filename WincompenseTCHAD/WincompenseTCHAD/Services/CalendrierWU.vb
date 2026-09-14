Option Strict On
Option Explicit On

Imports System.Data
Imports System.Data.SqlClient

''' <summary>
''' Jours ouvrés de la banque : ni samedi, ni dimanche, ni jour férié.
'''
''' La compensation porte l'activité du jour J en valeur au premier jour ouvré suivant. Une
''' écriture datée d'un jour chômé est rejetée par le core banking, ou repoussée d'office au jour
''' suivant sans que personne ne le sache — et la date de valeur ne correspond alors plus à ce
''' que la comptabilité attend.
'''
''' Les samedis et dimanches se déduisent du calendrier. Les jours fériés, non : ils changent
''' chaque année, et les fêtes musulmanes suivent le calendrier lunaire. Ils sont donc lus dans
''' T_JourFerieWU, que la banque complète sans recompiler l'application.
''' </summary>
Public NotInheritable Class CalendrierWU

    Private Sub New()
    End Sub

    Private Const TABLE As String = "T_JourFerieWU"
    Private Const ERREUR_TABLE_ABSENTE As Integer = 208

    ''' <summary>
    ''' Nombre de jours au-delà duquel on cesse de chercher un jour ouvré.
    ''' Une année entière de fériés consécutifs n'existe pas : passé ce cap, c'est que la table
    ''' contient une anomalie, et boucler sans fin serait pire que de rendre la main.
    ''' </summary>
    Private Const RECHERCHE_MAXIMALE As Integer = 370

    ''' <summary>
    ''' Jours fériés déjà lus, gardés le temps de l'exécution. La table ne change pas en cours
    ''' de journée, et la relire à chaque test de date multiplierait les allers-retours SQL.
    ''' </summary>
    Private Shared _feries As HashSet(Of Date)

    Private Shared _messageLecture As String = String.Empty

#Region "Jours ouvrés"

    ''' <summary>Vrai si la date est un jour ouvré pour la banque.</summary>
    Public Shared Function EstJourOuvre(jour As Date) As Boolean

        If jour.DayOfWeek = DayOfWeek.Saturday OrElse jour.DayOfWeek = DayOfWeek.Sunday Then
            Return False
        End If

        Return Not JoursFeries().Contains(jour.Date)
    End Function

    ''' <summary>
    ''' Premier jour ouvré STRICTEMENT postérieur à la date donnée.
    '''
    ''' C'est la date de valeur des écritures : l'activité du vendredi, du samedi et du dimanche
    ''' est donc portée au lundi — ou au mardi si le lundi est férié.
    ''' </summary>
    Public Shared Function ProchainJourOuvre(depuis As Date) As Date

        Dim candidat As Date = depuis.Date.AddDays(1)

        For essai As Integer = 1 To RECHERCHE_MAXIMALE
            If EstJourOuvre(candidat) Then Return candidat
            candidat = candidat.AddDays(1)
        Next

        ' Sortie de secours : plutôt rendre une date que boucler indéfiniment.
        Return depuis.Date.AddDays(1)
    End Function

    ''' <summary>
    ''' Explique en clair le passage de la date d'activité à la date de valeur, pour que
    ''' l'utilisateur voie pourquoi elle a été décalée.
    ''' </summary>
    Public Shared Function Explication(dateActivite As Date, dateValeur As Date) As String

        Dim ecart As Integer = CInt((dateValeur.Date - dateActivite.Date).TotalDays)

        If ecart <= 1 Then Return String.Empty

        Dim chomes As New List(Of String)
        Dim jour As Date = dateActivite.Date.AddDays(1)

        While jour < dateValeur.Date
            chomes.Add($"{NomDuJour(jour)} {jour:dd/MM}")
            jour = jour.AddDays(1)
        End While

        Return $"Report de {ecart} jours : {String.Join(", ", chomes)} " &
               If(chomes.Count > 1, "sont chômés.", "est chômé.")
    End Function

    Private Shared Function NomDuJour(jour As Date) As String

        Select Case jour.DayOfWeek
            Case DayOfWeek.Monday : Return "lundi"
            Case DayOfWeek.Tuesday : Return "mardi"
            Case DayOfWeek.Wednesday : Return "mercredi"
            Case DayOfWeek.Thursday : Return "jeudi"
            Case DayOfWeek.Friday : Return "vendredi"
            Case DayOfWeek.Saturday : Return "samedi"
            Case Else : Return "dimanche"
        End Select
    End Function

#End Region

#Region "Lecture des jours fériés"

    ''' <summary>Jours fériés enregistrés, lus une fois puis conservés.</summary>
    Private Shared Function JoursFeries() As HashSet(Of Date)

        If _feries IsNot Nothing Then Return _feries

        _feries = New HashSet(Of Date)()
        _messageLecture = String.Empty

        Const requete As String = "SELECT DateFerie FROM " & TABLE

        Try
            Using connexion As SqlConnection = WURepository.CreerConnexion()
                connexion.Open()

                Using commande As New SqlCommand(requete, connexion)
                    Using lecteur As SqlDataReader = commande.ExecuteReader()
                        While lecteur.Read()
                            If lecteur.IsDBNull(0) Then Continue While
                            _feries.Add(lecteur.GetDateTime(0).Date)
                        End While
                    End Using
                End Using
            End Using

        Catch ex As SqlException
            _messageLecture = If(ex.Number = ERREUR_TABLE_ABSENTE,
                                 "La table T_JourFerieWU n'existe pas : exécutez Scripts\10_JoursFeries.sql.",
                                 $"Lecture des jours fériés impossible : {ex.Message}")
        Catch ex As InvalidOperationException
            _messageLecture = $"Connexion SQL Server indisponible : {ex.Message}"
        End Try

        Return _feries
    End Function

    ''' <summary>
    ''' Avertissement à présenter avant de dater des écritures, ou chaîne vide si tout va bien.
    '''
    ''' Le silence serait la pire réponse : une année sans jour férié enregistré donne des dates
    ''' de valeur d'apparence normale, et l'erreur ne se découvrirait qu'au rejet du fichier.
    ''' </summary>
    Public Shared Function Avertissement(annee As Integer) As String

        Dim feries As HashSet(Of Date) = JoursFeries()

        If _messageLecture.Length > 0 Then
            Return _messageLecture & Environment.NewLine &
                   "Seuls les samedis et dimanches sont écartés : un jour férié passerait inaperçu."
        End If

        ' Any et non Count : sur un HashSet, Count est une propriété, et elle masque la méthode
        ' d'extension Count de LINQ — Visual Basic lit alors l'appel comme une indexation de la
        ' propriété. Any dit d'ailleurs mieux ce qu'on cherche, et s'arrête au premier trouvé.
        If feries.Any(Function(j) j.Year = annee) Then Return String.Empty

        Return $"Aucun jour férié n'est enregistré pour {annee}." & Environment.NewLine &
               "Seuls les samedis et dimanches sont écartés : complétez T_JourFerieWU, " &
               "notamment les fêtes musulmanes, qui ne se calculent pas d'avance."
    End Function

    ''' <summary>
    ''' Oblige à relire la table au prochain appel. À utiliser après avoir ajouté un jour férié
    ''' sans redémarrer l'application.
    ''' </summary>
    Public Shared Sub Oublier()
        _feries = Nothing
        _messageLecture = String.Empty
    End Sub

#End Region

End Class
