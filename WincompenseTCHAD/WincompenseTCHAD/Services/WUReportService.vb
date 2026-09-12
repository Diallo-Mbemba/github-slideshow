Option Strict On
Option Explicit On

Imports System.Data
Imports System.Globalization
Imports System.IO
Imports System.Text

''' <summary>
''' Exception métier levée lorsqu'un rapport Western Union est invalide
''' (colonne absente, fichier vide, incohérence de date, etc.).
''' Permet d'afficher un message explicite à l'utilisateur sans faire planter l'application.
''' </summary>
Public Class RapportInvalideException
    Inherits Exception

    Public Sub New(message As String)
        MyBase.New(message)
    End Sub

    Public Sub New(message As String, innerException As Exception)
        MyBase.New(message, innerException)
    End Sub
End Class

''' <summary>Résultat de l'agrégation du rapport d'activité pour un Account donné.</summary>
Public Class ActiviteAgregat
    Public Property PrincipalEnvoi As Decimal = 0D
    Public Property PrincipalPaye As Decimal = 0D
    Public Property ChargeEnvoi As Decimal = 0D
    Public Property Taxes As Decimal = 0D
End Class

''' <summary>Résultat de l'agrégation du rapport de règlement pour un Account donné.</summary>
Public Class ReglementAgregat
    Public Property CommissionPaiement As Decimal = 0D
End Class

''' <summary>
''' Service responsable de la lecture des fichiers plats (activité / règlement),
''' du parsing robuste des valeurs numériques et de l'agrégation par Account.
''' Ne contient aucune logique de calcul de commission (voir WUCalculationService).
''' </summary>
Public NotInheritable Class WUReportService

    Private Sub New()
    End Sub

#Region "Lecture de fichier"

    ''' <summary>
    ''' Lit un rapport Western Union au format texte séparé par tabulations (encodage Windows/ANSI)
    ''' et retourne son contenu sous forme de DataTable (première ligne = en-têtes).
    ''' </summary>
    ''' <param name="cheminFichier">Chemin complet du fichier à charger.</param>
    ''' <exception cref="RapportInvalideException">
    ''' Levée si le fichier est introuvable, verrouillé, invalide ou vide.
    ''' </exception>
    Public Shared Function LireRapportWU(cheminFichier As String) As DataTable

        If String.IsNullOrWhiteSpace(cheminFichier) Then
            Throw New RapportInvalideException("Le chemin du fichier est vide ou invalide.")
        End If

        If Not File.Exists(cheminFichier) Then
            Throw New RapportInvalideException($"Le fichier '{cheminFichier}' est introuvable.")
        End If

        Dim lignes As String()
        Try
            ' Encodage Windows/ANSI conforme aux exports Western Union.
            lignes = File.ReadAllLines(cheminFichier, Encoding.Default)
        Catch ex As IOException
            Throw New RapportInvalideException(
                $"Impossible de lire le fichier '{Path.GetFileName(cheminFichier)}'. " &
                "Il est peut-être ouvert dans une autre application (verrouillé).", ex)
        Catch ex As UnauthorizedAccessException
            Throw New RapportInvalideException(
                $"Accès refusé au fichier '{Path.GetFileName(cheminFichier)}'.", ex)
        End Try

        If lignes Is Nothing OrElse lignes.Length = 0 Then
            Throw New RapportInvalideException($"Le fichier '{Path.GetFileName(cheminFichier)}' est vide.")
        End If

        Dim table As New DataTable()
        Dim entetes As String() = lignes(0).Split(ControlChars.Tab)

        For Each entete As String In entetes
            Dim nomColonne As String = If(entete, String.Empty).Trim()
            If String.IsNullOrEmpty(nomColonne) Then
                nomColonne = "Colonne" & table.Columns.Count.ToString(CultureInfo.InvariantCulture)
            End If
            ' Évite les doublons de noms de colonnes (fichiers mal formés).
            Dim nomFinal As String = nomColonne
            Dim compteur As Integer = 1
            While table.Columns.Contains(nomFinal)
                nomFinal = nomColonne & "_" & compteur.ToString(CultureInfo.InvariantCulture)
                compteur += 1
            End While
            table.Columns.Add(nomFinal, GetType(String))
        Next

        If table.Columns.Count = 0 Then
            Throw New RapportInvalideException($"Le fichier '{Path.GetFileName(cheminFichier)}' ne contient aucune colonne exploitable.")
        End If

        For i As Integer = 1 To lignes.Length - 1
            Dim ligne As String = lignes(i)
            If String.IsNullOrWhiteSpace(ligne) Then
                Continue For ' ligne vide ignorée silencieusement
            End If

            Dim valeurs As String() = ligne.Split(ControlChars.Tab)
            Dim ligneDT As DataRow = table.NewRow()

            For c As Integer = 0 To table.Columns.Count - 1
                If c < valeurs.Length Then
                    ligneDT(c) = valeurs(c)
                Else
                    ' Ligne de rapport incomplète : on complète avec une valeur vide plutôt que de planter.
                    ligneDT(c) = String.Empty
                End If
            Next

            table.Rows.Add(ligneDT)
        Next

        Return table
    End Function

#End Region

#Region "Parsing robuste"

    ''' <summary>
    ''' Convertit une valeur quelconque (chaîne, DBNull, objet) en Decimal de façon robuste,
    ''' quel que soit le format du rapport Western Union :
    '''   - ancien format, séparateur décimal VIRGULE  : "1493,91" ;
    '''   - nouveau format, séparateur décimal POINT   : "-250000.0000".
    '''
    ''' Le séparateur décimal est déterminé par la position des symboles plutôt que par une
    ''' culture supposée : c'est TOUJOURS le dernier symbole rencontré ("." ou ","), l'éventuel
    ''' autre symbole étant alors un séparateur de milliers, supprimé avant l'analyse. Sans cette
    ''' précaution, une valeur comme "1.234" serait lue 1234 par la culture fr-FR (qui prend le
    ''' point pour un séparateur de milliers), soit une erreur d'un facteur 1000.
    '''
    ''' Retourne 0 (jamais d'exception) si la valeur ne peut pas être interprétée.
    ''' </summary>
    Public Shared Function ToDecimalSafe(value As Object) As Decimal

        If value Is Nothing OrElse value Is DBNull.Value Then
            Return 0D
        End If

        Dim texte As String = value.ToString()
        If String.IsNullOrWhiteSpace(texte) Then
            Return 0D
        End If

        texte = texte.Trim()
        ' Supprime les espaces (séparateurs de milliers éventuels) et les espaces insécables.
        texte = texte.Replace(" ", "").Replace(Convert.ToChar(&HA0).ToString(), "")

        Dim posPoint As Integer = texte.LastIndexOf("."c)
        Dim posVirgule As Integer = texte.LastIndexOf(","c)

        If posPoint >= 0 AndAlso posVirgule >= 0 Then
            ' Les deux symboles sont présents : le dernier est le séparateur décimal,
            ' le premier n'est qu'un séparateur de milliers et doit disparaître.
            If posVirgule > posPoint Then
                texte = texte.Replace(".", "").Replace(","c, "."c)
            Else
                texte = texte.Replace(",", "")
            End If
        ElseIf posVirgule >= 0 Then
            ' Seule la virgule est présente : c'est le séparateur décimal (ancien format).
            texte = texte.Replace(","c, "."c)
        End If
        ' Seul le point présent (nouveau format), ou aucun séparateur : rien à normaliser.

        Dim resultat As Decimal
        If Decimal.TryParse(texte, NumberStyles.Float, CultureInfo.InvariantCulture, resultat) Then
            Return resultat
        End If

        ' Valeur non interprétable : on ne plante pas, on retourne 0.
        Return 0D
    End Function

#End Region

#Region "Validation structurelle"

    ''' <summary>
    ''' Vérifie que toutes les colonnes obligatoires sont présentes dans le rapport.
    ''' </summary>
    ''' <param name="table">DataTable issue de LireRapportWU.</param>
    ''' <param name="colonnesRequises">Liste des noms de colonnes obligatoires.</param>
    ''' <param name="nomRapport">Nom du rapport pour un message d'erreur explicite.</param>
    ''' <exception cref="RapportInvalideException">Levée si une colonne obligatoire est absente.</exception>
    Public Shared Sub VerifierColonnesRapport(table As DataTable, colonnesRequises As IEnumerable(Of String), nomRapport As String)

        If table Is Nothing Then
            Throw New RapportInvalideException($"Le rapport {nomRapport} n'a pas pu être chargé.")
        End If

        For Each colonne As String In colonnesRequises
            If Not table.Columns.Contains(colonne) Then
                Throw New RapportInvalideException($"La colonne {colonne} est absente du rapport {nomRapport}.")
            End If
        Next
    End Sub

#End Region

#Region "Agrégation du rapport d'activité"

    ''' <summary>
    ''' Détermine si une ligne du rapport d'activité doit être incluse dans l'agrégation.
    ''' Sont exclues les lignes sans Account, ainsi que celles dont la colonne STATUS figure
    ''' dans ConstantesWU.StatutsActiviteExclus (aujourd'hui "C" = transaction ANNULÉE).
    ''' La colonne STATUS n'existe que dans le nouveau format de rapport Western Union : si elle
    ''' est absente (ancien format), aucune ligne n'est écartée à ce titre et le comportement
    ''' historique est conservé à l'identique.
    ''' </summary>
    Public Shared Function InclureLigneActivite(row As DataRow) As Boolean

        If row Is Nothing Then Return False

        If String.IsNullOrWhiteSpace(ObtenirValeurTexte(row, "Account")) Then
            Return False
        End If

        Dim statut As String = ObtenirValeurTexte(row, ConstantesWU.COLONNE_STATUS_ACTIVITE).Trim()
        If statut.Length = 0 Then
            Return True ' Colonne absente ou non renseignée : ligne conservée.
        End If

        For Each statutExclu As String In ConstantesWU.StatutsActiviteExclus
            If String.Equals(statut, statutExclu, StringComparison.OrdinalIgnoreCase) Then
                Return False
            End If
        Next

        Return True
    End Function

    ''' <summary>
    ''' Agrège le rapport d'activité par Account.
    ''' SendPayIndicator = "S" (Envoi) alimente PrincipalEnvoi / ChargeEnvoi / Taxes.
    ''' SendPayIndicator = "P" (Paiement) alimente PrincipalPaye.
    ''' Seules les lignes retenues par InclureLigneActivite sont agrégées.
    ''' Les valeurs sont conservées en Decimal, sans arrondi intermédiaire.
    ''' </summary>
    Public Shared Function CalculerActivite(dtActivite As DataTable) As Dictionary(Of String, ActiviteAgregat)

        Dim resultat As New Dictionary(Of String, ActiviteAgregat)(StringComparer.OrdinalIgnoreCase)

        If dtActivite Is Nothing Then Return resultat

        For Each row As DataRow In dtActivite.Rows

            ' Account vide ou transaction annulée : ligne ignorée, sans faire échouer le traitement.
            If Not InclureLigneActivite(row) Then
                Continue For
            End If

            Dim account As String = ObtenirValeurTexte(row, "Account")

            If Not resultat.ContainsKey(account) Then
                resultat(account) = New ActiviteAgregat()
            End If

            Dim agregat As ActiviteAgregat = resultat(account)
            Dim indicateur As String = ObtenirValeurTexte(row, "SendPayIndicator").Trim().ToUpperInvariant()

            Select Case indicateur
                Case "S"
                    agregat.PrincipalEnvoi += ToDecimalSafe(ObtenirValeur(row, "RecPrincipalREC"))
                    agregat.ChargeEnvoi += ToDecimalSafe(ObtenirValeur(row, "TotalChargesREC"))
                    agregat.Taxes += ToDecimalSafe(ObtenirValeur(row, "TaxesREC"))
                Case "P"
                    agregat.PrincipalPaye += ToDecimalSafe(ObtenirValeur(row, "PayPrincipalPAY"))
                Case Else
                    ' Indicateur inconnu ou vide : ligne non exploitable pour l'agrégation, ignorée sans erreur bloquante.
            End Select
        Next

        Return resultat
    End Function

#End Region

#Region "Agrégation du rapport de règlement"

    ''' <summary>
    ''' Détermine si une ligne du rapport de règlement doit être incluse dans les calculs.
    ''' RÈGLE PAR DÉFAUT (à affiner ultérieurement selon consigne métier) :
    ''' on reproduit le comportement du classeur bancaire existant et on inclut TOUTES les lignes,
    ''' y compris celles avec TransactionType = "A" (ajustements / reprocess), dès lors que
    ''' l'Account n'est pas vide. Aucune suppression arbitraire n'est effectuée.
    ''' À CONFIRMER : une règle métier spécifique pourra être ajoutée ici (ex. exclusion de
    ''' certains types d'ajustement) sans impacter le reste du service.
    ''' </summary>
    Public Shared Function InclureLigneReglement(row As DataRow) As Boolean

        If row Is Nothing Then Return False

        Dim account As String = ObtenirValeurTexte(row, "Account")
        If String.IsNullOrWhiteSpace(account) Then
            Return False
        End If

        ' TransactionType = "A" (ajustement/reprocess) : inclus par défaut, cf. commentaire ci-dessus.
        Return True
    End Function

    ''' <summary>
    ''' Détermine le facteur à appliquer aux montants LOC d'une ligne de règlement pour les
    ''' exprimer en FCFA, d'après la devise déclarée par la ligne (colonne LOCCurrencyCode) :
    '''   - XAF : montants déjà en FCFA, aucune conversion (facteur 1) — nouveau format WU ;
    '''   - EUR : conversion par TAUX_CONVERSION (655,957) — ancien format WU ;
    '''   - colonne absente : TAUX_CONVERSION, afin de reproduire à l'identique le comportement
    '''     historique sur un fichier qui ne déclarerait pas sa devise ;
    '''   - toute autre devise : facteur 1, les montants étant alors pris tels quels (cas non
    '''     rencontré à ce jour, la devise LOC étant celle du pays de l'agent). À surveiller.
    ''' </summary>
    Public Shared Function ObtenirFacteurConversion(row As DataRow) As Decimal

        If row Is Nothing OrElse Not row.Table.Columns.Contains(ConstantesWU.COLONNE_DEVISE_LOC) Then
            Return ConstantesWU.TAUX_CONVERSION
        End If

        Dim devise As String = ObtenirValeurTexte(row, ConstantesWU.COLONNE_DEVISE_LOC).Trim()

        If String.Equals(devise, ConstantesWU.DEVISE_EURO, StringComparison.OrdinalIgnoreCase) Then
            Return ConstantesWU.TAUX_CONVERSION
        End If

        Return 1D
    End Function

    ''' <summary>Indique si la ligne concerne un paiement effectué dans le pays local (Tchad).</summary>
    Private Shared Function EstPaiementLocal(payCountry As String) As Boolean
        For Each pays As String In ConstantesWU.PaysPaiementLocal
            If String.Equals(payCountry, pays, StringComparison.OrdinalIgnoreCase) Then
                Return True
            End If
        Next
        Return False
    End Function

    ''' <summary>
    ''' Agrège le rapport de règlement par Account pour calculer la commission de paiement locale.
    ''' CommissionPaiement += Abs(ClearChargesLOC + ClearFXLOC) * facteur de conversion, uniquement
    ''' pour les lignes payées au Tchad (PayCountry = "CHAD" ou "TCHAD" selon le format du rapport)
    ''' et retenues par InclureLigneReglement. Le facteur de conversion dépend de la devise déclarée
    ''' par la ligne (voir ObtenirFacteurConversion) : les montants LOC du nouveau format sont déjà
    ''' en FCFA et ne doivent donc plus être multipliés par 655,957.
    ''' </summary>
    Public Shared Function CalculerReglement(dtReglement As DataTable) As Dictionary(Of String, ReglementAgregat)

        Dim resultat As New Dictionary(Of String, ReglementAgregat)(StringComparer.OrdinalIgnoreCase)

        If dtReglement Is Nothing Then Return resultat

        For Each row As DataRow In dtReglement.Rows

            If Not InclureLigneReglement(row) Then
                Continue For
            End If

            Dim account As String = ObtenirValeurTexte(row, "Account")
            Dim payCountry As String = ObtenirValeurTexte(row, "PayCountry").Trim()

            If Not EstPaiementLocal(payCountry) Then
                Continue For
            End If

            If Not resultat.ContainsKey(account) Then
                resultat(account) = New ReglementAgregat()
            End If

            Dim clearCharges As Decimal = ToDecimalSafe(ObtenirValeur(row, "ClearChargesLOC"))
            Dim clearFX As Decimal = ToDecimalSafe(ObtenirValeur(row, "ClearFXLOC"))

            resultat(account).CommissionPaiement += Math.Abs(clearCharges + clearFX) * ObtenirFacteurConversion(row)
        Next

        Return resultat
    End Function

#End Region

#Region "Validation des dates (section 16)"

    ''' <summary>
    ''' Extrait, si possible, la date unique du rapport d'activité (colonne txnDateLOC).
    ''' Retourne Nothing si la colonne est absente ou si aucune date exploitable n'est trouvée.
    ''' </summary>
    Public Shared Function ObtenirDateRapport(table As DataTable, nomColonneDate As String) As Date?

        If table Is Nothing OrElse Not table.Columns.Contains(nomColonneDate) Then
            Return Nothing
        End If

        For Each row As DataRow In table.Rows
            Dim texte As String = ObtenirValeurTexte(row, nomColonneDate)
            If String.IsNullOrWhiteSpace(texte) Then Continue For

            Dim dateValeur As Date

            ' Formats produits par Western Union, essayés en premier : "20260530" (yyyyMMdd) n'est
            ' reconnu par aucune culture et échouerait silencieusement avec un simple TryParse.
            If Date.TryParseExact(texte, ConstantesWU.FormatsDateRapport, CultureInfo.InvariantCulture,
                                  DateTimeStyles.None, dateValeur) Then
                Return dateValeur.Date
            End If

            If Date.TryParse(texte, CultureInfo.CurrentCulture, DateTimeStyles.None, dateValeur) OrElse
               Date.TryParse(texte, CultureInfo.InvariantCulture, DateTimeStyles.None, dateValeur) Then
                Return dateValeur.Date
            End If
        Next

        Return Nothing
    End Function

    ''' <summary>
    ''' Reconstitue une date à partir d'un triplet de colonnes Année/Mois/Jour préfixées
    ''' (ex. SetDateLOCYear / SetDateLOCMonth / SetDateLOCDay). Retourne Nothing si les trois
    ''' colonnes ne sont pas toutes présentes, ou si aucune ligne ne porte une date valide
    ''' (les rapports contiennent des triplets neutres "0/0/0" pour les transactions non encore
    ''' payées : ces lignes sont ignorées).
    ''' </summary>
    Public Shared Function ObtenirDateDepuisComposants(table As DataTable, prefixe As String) As Date?

        If table Is Nothing Then Return Nothing

        Dim colAnnee As String = prefixe & ConstantesWU.SUFFIXE_DATE_ANNEE
        Dim colMois As String = prefixe & ConstantesWU.SUFFIXE_DATE_MOIS
        Dim colJour As String = prefixe & ConstantesWU.SUFFIXE_DATE_JOUR

        If Not (table.Columns.Contains(colAnnee) AndAlso
                table.Columns.Contains(colMois) AndAlso
                table.Columns.Contains(colJour)) Then
            Return Nothing
        End If

        For Each row As DataRow In table.Rows
            Dim annee, mois, jour As Integer
            If Integer.TryParse(ObtenirValeurTexte(row, colAnnee), annee) AndAlso
               Integer.TryParse(ObtenirValeurTexte(row, colMois), mois) AndAlso
               Integer.TryParse(ObtenirValeurTexte(row, colJour), jour) Then

                If annee > 0 AndAlso mois >= 1 AndAlso mois <= 12 AndAlso jour >= 1 AndAlso jour <= 31 Then
                    Try
                        Return New Date(annee, mois, jour)
                    Catch ex As ArgumentOutOfRangeException
                        ' Triplet incohérent (ex. 31 février) : on poursuit avec les lignes suivantes.
                    End Try
                End If
            End If
        Next

        Return Nothing
    End Function

    ''' <summary>
    ''' Détermine la date du rapport d'activité : colonne de date simple (txnDateLOC) en priorité,
    ''' puis reconstitution depuis les colonnes Année/Mois/Jour. Retourne Nothing si aucune date
    ''' exploitable n'a pu être trouvée.
    ''' </summary>
    Public Shared Function ObtenirDateActivite(table As DataTable) As Date?
        Dim resultat As Date? = ObtenirDateRapport(table, ConstantesWU.COLONNE_DATE_ACTIVITE)
        If resultat.HasValue Then Return resultat

        For Each prefixe As String In ConstantesWU.PrefixesDateActivite
            resultat = ObtenirDateDepuisComposants(table, prefixe)
            If resultat.HasValue Then Return resultat
        Next

        Return Nothing
    End Function

    ''' <summary>
    ''' Détermine la date du rapport de règlement. Ce rapport ne comportant pas de colonne de date
    ''' simple, la date est reconstituée depuis les colonnes Année/Mois/Jour de la date de règlement
    ''' locale (SetDateLOC), puis à défaut de la date d'édition du rapport (RepDate).
    ''' </summary>
    Public Shared Function ObtenirDateReglement(table As DataTable) As Date?
        Dim resultat As Date? = ObtenirDateRapport(table, ConstantesWU.COLONNE_DATE_REGLEMENT)
        If resultat.HasValue Then Return resultat

        For Each prefixe As String In ConstantesWU.PrefixesDateReglement
            resultat = ObtenirDateDepuisComposants(table, prefixe)
            If resultat.HasValue Then Return resultat
        Next

        Return Nothing
    End Function

    ''' <summary>
    ''' Vérifie que le rapport d'activité et le rapport de règlement correspondent à la même journée.
    ''' Si la colonne de date n'existe pas côté règlement, la vérification est ignorée avec un avertissement
    ''' (le traitement n'est pas bloqué dans ce cas, faute d'information comparable).
    ''' </summary>
    ''' <param name="messageAvertissement">Message informatif ou d'anomalie à afficher à l'utilisateur.</param>
    ''' <returns>False uniquement si une incohérence de date est avérée entre les deux rapports.</returns>
    Public Shared Function ValiderCoherenceDates(dtActivite As DataTable, dtReglement As DataTable, ByRef messageAvertissement As String) As Boolean

        messageAvertissement = String.Empty

        Dim dateActivite As Date? = ObtenirDateActivite(dtActivite)

        If dateActivite Is Nothing Then
            messageAvertissement = "Impossible de déterminer la date du rapport d'activité " &
                                    "(colonne txnDateLOC et colonnes Année/Mois/Jour absentes ou illisibles)."
            Return True ' Non bloquant : on ne peut simplement pas comparer.
        End If

        Dim dateReglement As Date? = ObtenirDateReglement(dtReglement)

        If dateReglement Is Nothing Then
            messageAvertissement = $"Date d'activité détectée : {dateActivite.Value:dd/MM/yyyy}. " &
                                    "Aucune colonne de date comparable trouvée dans le rapport de règlement " &
                                    "(vérification de cohérence ignorée)."
            Return True
        End If

        If dateActivite.Value.Date <> dateReglement.Value.Date Then
            messageAvertissement = $"Incohérence de date : rapport d'activité = {dateActivite.Value:dd/MM/yyyy}, " &
                                    $"rapport de règlement = {dateReglement.Value:dd/MM/yyyy}. Traitement bloqué."
            Return False
        End If

        messageAvertissement = $"Date de traitement : {dateActivite.Value:dd/MM/yyyy}."
        Return True
    End Function

#End Region

#Region "Utilitaires internes"

    Private Shared Function ObtenirValeur(row As DataRow, colonne As String) As Object
        If row Is Nothing OrElse Not row.Table.Columns.Contains(colonne) Then
            Return Nothing
        End If
        Dim valeur As Object = row(colonne)
        Return If(valeur Is DBNull.Value, Nothing, valeur)
    End Function

    Private Shared Function ObtenirValeurTexte(row As DataRow, colonne As String) As String
        Dim valeur As Object = ObtenirValeur(row, colonne)
        Return If(valeur Is Nothing, String.Empty, valeur.ToString())
    End Function

#End Region

End Class
