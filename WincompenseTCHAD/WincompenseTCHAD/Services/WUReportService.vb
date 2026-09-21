Option Strict On
Option Explicit On

Imports System.Data
Imports System.Globalization
Imports System.IO
Imports System.IO.Compression
Imports System.Linq
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

    ''' <summary>Nombre de transactions d'envoi retenues (SendPayIndicator = "S").</summary>
    Public Property NombreEnvois As Integer = 0

    ''' <summary>Nombre de transactions de paiement retenues (SendPayIndicator = "P").</summary>
    Public Property NombrePaiements As Integer = 0

    ''' <summary>
    ''' Nombre de transactions écartées de l'agrégation parce qu'annulées (STATUS = "C").
    ''' Sans effet sur les montants : compté pour figurer, à titre d'information, dans les
    ''' rapports d'activité — une journée riche en annulations mérite d'être regardée.
    ''' </summary>
    Public Property NombreAnnulations As Integer = 0

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
    ''' Lit un rapport Western Union et retourne son contenu sous forme de DataTable
    ''' (première ligne = en-têtes, colonnes séparées par des tabulations, encodage Windows/ANSI).
    '''
    ''' Le fichier peut être fourni indifféremment sous deux formes :
    '''   - une ARCHIVE ZIP contenant le rapport (forme livrée par Western Union) : le rapport en
    '''     est extrait et lu directement en mémoire, sans écrire de fichier temporaire ;
    '''   - le fichier TEXTE déjà décompressé.
    ''' La distinction se fait sur le contenu réel du fichier (signature "PK" des archives ZIP) et
    ''' non sur son extension, car les archives livrées portent le nom du rapport qu'elles
    ''' contiennent et peuvent donc prêter à confusion.
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
            lignes = LireLignesRapport(cheminFichier)
        Catch ex As InvalidDataException
            Throw New RapportInvalideException(
                $"L'archive '{Path.GetFileName(cheminFichier)}' est illisible ou endommagée.", ex)
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

    ''' <summary>
    ''' Retourne les lignes du rapport, que le fichier fourni soit une archive ZIP ou le fichier
    ''' texte lui-même. Dans le cas d'une archive, le rapport est lu directement depuis le flux
    ''' compressé : aucun fichier temporaire n'est créé, donc rien à nettoyer ensuite.
    ''' </summary>
    Private Shared Function LireLignesRapport(cheminFichier As String) As String()

        If Not EstArchiveZip(cheminFichier) Then
            ' Encodage Windows/ANSI conforme aux exports Western Union.
            Return File.ReadAllLines(cheminFichier, Encoding.Default)
        End If

        Using fluxArchive As FileStream = File.OpenRead(cheminFichier)
            Using archive As New ZipArchive(fluxArchive, ZipArchiveMode.Read)

                Dim entree As ZipArchiveEntry = ChoisirEntreeRapport(archive)
                If entree Is Nothing Then
                    Throw New RapportInvalideException(
                        $"L'archive '{Path.GetFileName(cheminFichier)}' ne contient aucun fichier exploitable.")
                End If

                Dim lignes As New List(Of String)
                Using lecteur As New StreamReader(entree.Open(), Encoding.Default)
                    While Not lecteur.EndOfStream
                        lignes.Add(lecteur.ReadLine())
                    End While
                End Using

                Return lignes.ToArray()
            End Using
        End Using
    End Function

    ''' <summary>
    ''' Indique si le fichier est une archive ZIP, d'après sa signature binaire ("PK", soit les
    ''' octets 0x50 0x4B) et non d'après son extension : les archives livrées par Western Union
    ''' portent le nom du rapport qu'elles contiennent, l'extension n'est donc pas fiable.
    ''' Retourne False sur toute erreur de lecture : le fichier sera alors traité comme du texte
    ''' et l'erreur remontera, plus explicite, lors de la lecture proprement dite.
    ''' </summary>
    Private Shared Function EstArchiveZip(cheminFichier As String) As Boolean

        Try
            Using flux As FileStream = File.OpenRead(cheminFichier)
                If flux.Length < 4 Then Return False

                Dim signature(1) As Byte
                If flux.Read(signature, 0, 2) < 2 Then Return False

                Return signature(0) = &H50 AndAlso signature(1) = &H4B
            End Using
        Catch
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Choisit, parmi les entrées d'une archive, celle qui constitue le rapport : le fichier .txt
    ''' le plus volumineux, ou à défaut le fichier le plus volumineux quelle que soit son extension.
    ''' Les répertoires (entrées sans nom de fichier) sont ignorés. Ce choix par la taille évite de
    ''' retenir par erreur un éventuel fichier annexe (accusé de réception, notice) joint au rapport.
    ''' </summary>
    Private Shared Function ChoisirEntreeRapport(archive As ZipArchive) As ZipArchiveEntry

        Dim fichiers As List(Of ZipArchiveEntry) =
            archive.Entries.Where(Function(e) Not String.IsNullOrEmpty(e.Name)).ToList()

        If fichiers.Count = 0 Then Return Nothing

        Dim textes As List(Of ZipArchiveEntry) =
            fichiers.Where(Function(e) e.Name.EndsWith(".txt", StringComparison.OrdinalIgnoreCase)).ToList()

        If textes.Count > 0 Then
            Return textes.OrderByDescending(Function(e) e.Length).First()
        End If

        Return fichiers.OrderByDescending(Function(e) e.Length).First()
    End Function


    ''' <summary>
    ''' Retourne le nom du rapport CONTENU dans l'archive, ou une chaîne vide si le fichier
    ''' n'est pas une archive (ou s'il est illisible). Sert aux contrôles de sécurité : quand
    ''' l'archive a été renommée, le fichier qu'elle contient porte encore, lui, la nomenclature
    ''' Western Union et permet donc d'identifier le type et la période du rapport.
    ''' Ne lève jamais d'exception : un nom indéterminé n'est pas une erreur de traitement.
    ''' </summary>
    Public Shared Function ObtenirNomRapportInterne(cheminFichier As String) As String

        If String.IsNullOrWhiteSpace(cheminFichier) OrElse Not File.Exists(cheminFichier) Then
            Return String.Empty
        End If

        If Not EstArchiveZip(cheminFichier) Then
            Return String.Empty
        End If

        Try
            Using fluxArchive As FileStream = File.OpenRead(cheminFichier)
                Using archive As New ZipArchive(fluxArchive, ZipArchiveMode.Read)
                    Dim entree As ZipArchiveEntry = ChoisirEntreeRapport(archive)
                    Return If(entree Is Nothing, String.Empty, entree.Name)
                End Using
            End Using
        Catch ex As InvalidDataException
            Return String.Empty
        Catch ex As IOException
            Return String.Empty
        Catch ex As UnauthorizedAccessException
            Return String.Empty
        End Try
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

            Dim account As String = ObtenirValeurTexte(row, "Account")

            ' Sans Account, la ligne n'est rattachable à aucun point de vente : rien à en faire.
            If String.IsNullOrWhiteSpace(account) Then
                Continue For
            End If

            If Not resultat.ContainsKey(account) Then
                resultat(account) = New ActiviteAgregat()
            End If

            Dim agregat As ActiviteAgregat = resultat(account)

            ' Transaction annulée : exclue des montants, mais comptée — le nombre d'annulations
            ' d'une journée est une information de gestion, pas un détail à faire disparaître.
            If Not InclureLigneActivite(row) Then
                agregat.NombreAnnulations += 1
                Continue For
            End If
            Dim indicateur As String = ObtenirValeurTexte(row, "SendPayIndicator").Trim().ToUpperInvariant()

            Select Case indicateur
                Case "S"
                    agregat.NombreEnvois += 1
                    agregat.PrincipalEnvoi += ToDecimalSafe(ObtenirValeur(row, "RecPrincipalREC"))
                    agregat.ChargeEnvoi += ToDecimalSafe(ObtenirValeur(row, "TotalChargesREC"))
                    agregat.Taxes += ToDecimalSafe(ObtenirValeur(row, "TaxesREC"))
                Case "P"
                    agregat.NombrePaiements += 1
                    agregat.PrincipalPaye += ToDecimalSafe(ObtenirValeur(row, "PayPrincipalPAY"))
                Case Else
                    ' Indicateur inconnu ou vide : ligne non exploitable pour l'agrégation, ignorée sans erreur bloquante.
            End Select
        Next

        Return resultat
    End Function

    ''' <summary>
    ''' Extrait du rapport d'activité le détail des transactions, une par ligne, identifiées par
    ''' leur MTCN.
    '''
    ''' Contrairement à CalculerActivite, qui agrège et écarte les annulations des montants, la
    ''' liste retournée CONSERVE toutes les lignes, annulations comprises : c'est précisément
    ''' pour pouvoir retrouver une transaction, quelle qu'ait été son issue, que ce détail est
    ''' historisé.
    '''
    ''' Les lignes sans Account ou sans MTCN sont ignorées : elles ne seraient rattachables à
    ''' rien et ne se retrouveraient pas.
    ''' </summary>
    ''' <param name="jour">Journée d'activité, reportée sur chaque transaction.</param>
    Public Shared Function ExtraireTransactions(dtActivite As DataTable, jour As Date) As List(Of TransactionWU)

        Dim resultat As New List(Of TransactionWU)

        If dtActivite Is Nothing Then Return resultat

        For Each row As DataRow In dtActivite.Rows

            Dim account As String = ObtenirValeurTexte(row, "Account").Trim()
            Dim mtcn As String = ObtenirValeurTexte(row, "MTCN").Trim()

            If account.Length = 0 OrElse mtcn.Length = 0 Then Continue For

            Dim indicateur As String = ObtenirValeurTexte(row, "SendPayIndicator").Trim().ToUpperInvariant()

            Dim sens As String
            Dim montant As Decimal

            Select Case indicateur
                Case "S"
                    sens = TransactionWU.SENS_ENVOI
                    montant = ToDecimalSafe(ObtenirValeur(row, "RecPrincipalREC"))
                Case "P"
                    sens = TransactionWU.SENS_PAIEMENT
                    montant = ToDecimalSafe(ObtenirValeur(row, "PayPrincipalPAY"))
                Case Else
                    ' Indicateur inconnu : la ligne n'est ni un envoi ni un paiement, rien à tracer.
                    Continue For
            End Select

            resultat.Add(New TransactionWU() With {
                .DateActivite = jour.Date,
                .Account = account,
                .MTCN = mtcn,
                .Sens = sens,
                .Statut = ObtenirValeurTexte(row, ConstantesWU.COLONNE_STATUS_ACTIVITE).Trim(),
                .Montant = montant
            })
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

            Dim dateValeur As Date
            If EssayerLireDate(ObtenirValeurTexte(row, nomColonneDate), dateValeur) Then
                Return dateValeur
            End If
        Next

        Return Nothing
    End Function

    ''' <summary>
    ''' Lit une date telle que Western Union l'écrit. Rend Faux si le texte n'en porte pas.
    '''
    ''' Isolée pour que la lecture d'UNE date et l'inventaire de TOUTES les journées du rapport
    ''' ne puissent pas comprendre deux choses différentes du même fichier.
    ''' </summary>
    Private Shared Function EssayerLireDate(texte As String, ByRef resultat As Date) As Boolean

        resultat = Date.MinValue
        If String.IsNullOrWhiteSpace(texte) Then Return False

        Dim lue As Date

        ' Formats produits par Western Union, essayés en premier : "20260530" (yyyyMMdd) n'est
        ' reconnu par aucune culture et échouerait silencieusement avec un simple TryParse.
        If Date.TryParseExact(texte, ConstantesWU.FormatsDateRapport, CultureInfo.InvariantCulture,
                              DateTimeStyles.None, lue) Then
            resultat = lue.Date
            Return True
        End If

        If Date.TryParse(texte, CultureInfo.CurrentCulture, DateTimeStyles.None, lue) OrElse
           Date.TryParse(texte, CultureInfo.InvariantCulture, DateTimeStyles.None, lue) Then
            resultat = lue.Date
            Return True
        End If

        Return False
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
    ''' Toutes les journées présentes dans le rapport d'activité, dédoublonnées et triées.
    '''
    ''' POURQUOI CETTE FONCTION EXISTE. ObtenirDateActivite rend la date de la PREMIÈRE ligne
    ''' exploitable, et rien ne vérifiait que les suivantes portaient le même jour. Un rapport
    ''' hebdomadaire — c'est ainsi que la banque liquide, « DU 08 AU 14/09 » — était donc agrégé
    ''' en entier puis étiqueté d'une seule journée : celle de sa première ligne. Les montants
    ''' étaient justes, l'intitulé faux, et rien ne le disait.
    '''
    ''' Rend une liste vide si aucune date n'est lisible : l'appelant décide, comme partout
    ''' ailleurs, plutôt que de se voir imposer une journée inventée.
    ''' </summary>
    Public Shared Function JourneesDuRapport(table As DataTable) As List(Of Date)

        Dim journees As New SortedSet(Of Date)()
        If table Is Nothing Then Return journees.ToList()

        ' Colonne de date simple d'abord, puis reconstitution depuis Année/Mois/Jour : le même
        ' ordre de préférence qu'ObtenirDateActivite, pour que les deux ne puissent pas lire
        ' deux choses différentes du même fichier.
        If table.Columns.Contains(ConstantesWU.COLONNE_DATE_ACTIVITE) Then

            For Each ligne As DataRow In table.Rows
                Dim lue As Date
                If EssayerLireDate(ObtenirValeurTexte(ligne, ConstantesWU.COLONNE_DATE_ACTIVITE), lue) Then
                    journees.Add(lue.Date)
                End If
            Next

            If journees.Count > 0 Then Return journees.ToList()
        End If

        For Each prefixe As String In ConstantesWU.PrefixesDateActivite

            Dim colAnnee As String = prefixe & ConstantesWU.SUFFIXE_DATE_ANNEE
            Dim colMois As String = prefixe & ConstantesWU.SUFFIXE_DATE_MOIS
            Dim colJour As String = prefixe & ConstantesWU.SUFFIXE_DATE_JOUR

            If Not (table.Columns.Contains(colAnnee) AndAlso
                    table.Columns.Contains(colMois) AndAlso
                    table.Columns.Contains(colJour)) Then Continue For

            For Each ligne As DataRow In table.Rows

                Dim annee, mois, jour As Integer
                If Not (Integer.TryParse(ObtenirValeurTexte(ligne, colAnnee), annee) AndAlso
                        Integer.TryParse(ObtenirValeurTexte(ligne, colMois), mois) AndAlso
                        Integer.TryParse(ObtenirValeurTexte(ligne, colJour), jour)) Then Continue For

                If annee <= 0 OrElse mois < 1 OrElse mois > 12 OrElse jour < 1 OrElse jour > 31 Then Continue For

                Try
                    journees.Add(New Date(annee, mois, jour))
                Catch ex As ArgumentOutOfRangeException
                    ' Triplet incohérent (31 février) : la ligne est ignorée, pas le rapport.
                End Try
            Next

            If journees.Count > 0 Then Return journees.ToList()
        Next

        Return journees.ToList()
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
