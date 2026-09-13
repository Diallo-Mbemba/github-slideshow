Option Strict On
Option Explicit On

Imports System.Data
Imports System.Globalization
Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions

''' <summary>
''' Nature d'un rapport Western Union.
''' </summary>
Public Enum TypeRapportWU
    ''' <summary>Type non déterminable (nom de fichier non standard, contenu non reconnu).</summary>
    Indetermine = 0
    ''' <summary>Rapport d'activité (bouton 1).</summary>
    Activite = 1
    ''' <summary>Rapport de règlement (bouton 2).</summary>
    Reglement = 2
End Enum

''' <summary>
''' Période couverte par un rapport, telle que déclarée dans le nom du fichier.
''' Un rapport quotidien porte une période d'un seul jour (Debut = Fin).
''' </summary>
Public Class PeriodeRapport

    Public Property Debut As Date
    Public Property Fin As Date

    Public Sub New(debutPeriode As Date, finPeriode As Date)
        ' Ordonne les bornes : le nom du fichier ne garantit pas leur ordre.
        If debutPeriode <= finPeriode Then
            Debut = debutPeriode.Date
            Fin = finPeriode.Date
        Else
            Debut = finPeriode.Date
            Fin = debutPeriode.Date
        End If
    End Sub

    ''' <summary>Vrai si la période ne couvre qu'une seule journée (cas normal du traitement J+1).</summary>
    Public ReadOnly Property EstJourUnique As Boolean
        Get
            Return Debut = Fin
        End Get
    End Property

    ''' <summary>Égalité stricte des deux bornes : deux rapports doivent couvrir exactement la même période.</summary>
    Public Function EstIdentiqueA(autre As PeriodeRapport) As Boolean
        If autre Is Nothing Then Return False
        Return Debut = autre.Debut AndAlso Fin = autre.Fin
    End Function

    Public Overrides Function ToString() As String
        If EstJourUnique Then
            Return Debut.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture)
        End If
        Return String.Format(CultureInfo.InvariantCulture, "du {0:dd/MM/yyyy} au {1:dd/MM/yyyy}", Debut, Fin)
    End Function

End Class

''' <summary>
''' Ce que l'on sait d'un fichier sélectionné par l'utilisateur, d'après son seul nom.
''' </summary>
Public Class InfosFichierRapport

    Public Property Chemin As String = String.Empty
    Public Property NomFichier As String = String.Empty

    ''' <summary>Type déduit du nom du fichier. Indetermine si le nom ne permet pas de conclure.</summary>
    Public Property Type As TypeRapportWU = TypeRapportWU.Indetermine

    ''' <summary>Période déduite du nom du fichier. Nothing si aucune date n'y est lisible.</summary>
    Public Property Periode As PeriodeRapport = Nothing

    ''' <summary>Nom réellement analysé (nom de l'archive, ou nom du rapport qu'elle contient).</summary>
    Public Property NomAnalyse As String = String.Empty

End Class

''' <summary>
''' Contrôles de sécurité sur les fichiers chargés, AVANT tout calcul comptable.
'''
''' Deux risques sont couverts, chacun par deux barrières indépendantes :
'''
'''   1. Charger un rapport d'ACTIVITÉ à la place d'un rapport de RÈGLEMENT (ou l'inverse).
'''      - barrière 1 (immédiate, au clic)  : nom du fichier (VerifierTypeAttendu) ;
'''      - barrière 2 (au traitement, FAIT FOI) : colonnes réellement présentes dans le fichier
'''        (VerifierTypeRapport). Un fichier renommé franchit la première, jamais la seconde.
'''
'''   2. Croiser deux rapports de PÉRIODES DIFFÉRENTES.
'''      - barrière 1 (immédiate, au clic)  : périodes lues dans les deux noms de fichiers
'''        (VerifierMemePeriode) ;
'''      - barrière 2 (au traitement, FAIT FOI) : dates lues DANS les rapports
'''        (WUReportService.ValiderCoherenceDates, déjà en place).
'''
''' Principe retenu : un contrôle sur le nom ne bloque QUE s'il constate une contradiction
''' certaine. Un nom illisible ou non standard n'interrompt jamais le traitement — il serait
''' inacceptable qu'un simple renommage empêche de comptabiliser la journée — mais le contrôle
''' sur le contenu, lui, reste systématique et bloquant.
''' </summary>
Public NotInheritable Class WUFichierService

    Private Sub New()
    End Sub

#Region "Analyse du nom de fichier"

    ''' <summary>
    ''' Analyse le nom du fichier sélectionné pour en déduire le type de rapport et la période.
    ''' Si le nom de l'archive ne permet pas de conclure (archive renommée, nom générique),
    ''' le nom du rapport CONTENU dans l'archive est analysé à son tour.
    ''' Cette fonction ne lève jamais d'exception : elle renvoie ce qu'elle a pu déterminer.
    ''' </summary>
    Public Shared Function AnalyserFichier(cheminFichier As String) As InfosFichierRapport

        Dim infos As New InfosFichierRapport()

        If String.IsNullOrWhiteSpace(cheminFichier) Then
            Return infos
        End If

        infos.Chemin = cheminFichier

        Try
            infos.NomFichier = Path.GetFileName(cheminFichier)
        Catch ex As ArgumentException
            ' Chemin contenant des caractères invalides : on travaille sur la chaîne brute.
            infos.NomFichier = cheminFichier
        End Try

        infos.NomAnalyse = infos.NomFichier
        infos.Type = DetecterTypeDepuisNom(infos.NomFichier)
        infos.Periode = ExtrairePeriode(infos.NomFichier)

        ' Le nom de l'archive n'est pas concluant : on se rabat sur celui du rapport qu'elle
        ' contient, qui porte la nomenclature Western Union même si l'archive a été renommée.
        If infos.Type = TypeRapportWU.Indetermine OrElse infos.Periode Is Nothing Then

            Dim nomInterne As String = WUReportService.ObtenirNomRapportInterne(cheminFichier)

            If Not String.IsNullOrWhiteSpace(nomInterne) Then

                If infos.Type = TypeRapportWU.Indetermine Then
                    Dim typeInterne As TypeRapportWU = DetecterTypeDepuisNom(nomInterne)
                    If typeInterne <> TypeRapportWU.Indetermine Then
                        infos.Type = typeInterne
                        infos.NomAnalyse = nomInterne
                    End If
                End If

                If infos.Periode Is Nothing Then
                    Dim periodeInterne As PeriodeRapport = ExtrairePeriode(nomInterne)
                    If periodeInterne IsNot Nothing Then
                        infos.Periode = periodeInterne
                        infos.NomAnalyse = nomInterne
                    End If
                End If
            End If
        End If

        Return infos
    End Function

    ''' <summary>
    ''' Déduit le type de rapport des mots-clés présents dans le nom du fichier.
    ''' Si le nom contient à la fois un marqueur d'activité et un marqueur de règlement,
    ''' aucune conclusion n'est tirée (Indetermine) : mieux vaut laisser le contrôle sur le
    ''' contenu trancher que de se tromper sur un nom ambigu.
    ''' </summary>
    Public Shared Function DetecterTypeDepuisNom(nomFichier As String) As TypeRapportWU

        If String.IsNullOrWhiteSpace(nomFichier) Then
            Return TypeRapportWU.Indetermine
        End If

        Dim nom As String = Normaliser(nomFichier)

        Dim estActivite As Boolean = ContientUnMarqueur(nom, ConstantesWU.MarqueursNomActivite)
        Dim estReglement As Boolean = ContientUnMarqueur(nom, ConstantesWU.MarqueursNomReglement)

        If estActivite AndAlso Not estReglement Then Return TypeRapportWU.Activite
        If estReglement AndAlso Not estActivite Then Return TypeRapportWU.Reglement

        Return TypeRapportWU.Indetermine
    End Function

    ''' <summary>
    ''' Extrait la période couverte par le rapport du nom du fichier. Deux nomenclatures :
    '''
    '''   - Western Union : "RSP_TD383_ACTIVITY_REPORT_BY_ACCOUNT_20260530_20260530_202606031151"
    '''     soit deux dates yyyyMMdd (début et fin de période) suivies de l'horodatage d'édition.
    '''     Seuls les groupes de 8 chiffres EXACTEMENT sont retenus : l'horodatage d'édition, qui
    '''     en compte 12, est ainsi écarté et ne peut pas être pris pour une date de période.
    '''
    '''   - anciens rapports : "Rapport d'activité par Site ... du 02 Jan 2021", soit une date
    '''     unique sous forme jour / mois nommé / année.
    '''
    ''' Retourne Nothing si aucune date exploitable n'est trouvée.
    ''' </summary>
    Public Shared Function ExtrairePeriode(nomFichier As String) As PeriodeRapport

        If String.IsNullOrWhiteSpace(nomFichier) Then
            Return Nothing
        End If

        Dim nom As String = Normaliser(nomFichier)

        ' --- Nomenclature Western Union : groupes de 8 chiffres isolés (yyyyMMdd) ---
        Dim dates As New List(Of Date)

        For Each correspondance As Match In Regex.Matches(nom, "(?<!\d)\d{8}(?!\d)")
            Dim valeur As Date
            If Date.TryParseExact(correspondance.Value, "yyyyMMdd", CultureInfo.InvariantCulture,
                                  DateTimeStyles.None, valeur) Then
                dates.Add(valeur.Date)
            End If
        Next

        If dates.Count > 0 Then
            Dim minimum As Date = dates(0)
            Dim maximum As Date = dates(0)
            For Each valeur As Date In dates
                If valeur < minimum Then minimum = valeur
                If valeur > maximum Then maximum = valeur
            Next
            Return New PeriodeRapport(minimum, maximum)
        End If

        ' --- Anciens rapports : "du 02 Jan 2021" (séparateurs espace, tiret ou souligné) ---
        Dim motif As String = "(?<!\d)(\d{1,2})[\s_\-\.]+([a-z]{3,10})[\s_\-\.]+((?:19|20)\d{2})(?!\d)"
        For Each correspondance As Match In Regex.Matches(nom, motif)

            Dim jour As Integer
            Dim annee As Integer
            If Not Integer.TryParse(correspondance.Groups(1).Value, jour) Then Continue For
            If Not Integer.TryParse(correspondance.Groups(3).Value, annee) Then Continue For

            Dim mois As Integer = NumeroDuMois(correspondance.Groups(2).Value)
            If mois = 0 Then Continue For

            Try
                Dim valeur As New Date(annee, mois, jour)
                Return New PeriodeRapport(valeur, valeur)
            Catch ex As ArgumentOutOfRangeException
                ' Date impossible (ex. 31 février) : on poursuit la recherche.
            End Try
        Next

        Return Nothing
    End Function

    ''' <summary>
    ''' Retourne le numéro du mois désigné par un libellé français ou anglais, abrégé ou complet,
    ''' ou 0 si le libellé n'est pas reconnu. Un libellé ambigu comme "jui" (juin ? juillet ?)
    ''' n'est volontairement PAS reconnu : sur une donnée comptable, ne rien conclure vaut mieux
    ''' que deviner.
    ''' </summary>
    Public Shared Function NumeroDuMois(libelle As String) As Integer

        If String.IsNullOrWhiteSpace(libelle) Then Return 0

        Dim texte As String = Normaliser(libelle)

        For i As Integer = 0 To ConstantesWU.MoisNommes.Length - 1
            If String.Equals(texte, ConstantesWU.MoisNommes(i), StringComparison.Ordinal) Then
                Return ConstantesWU.NumerosMoisNommes(i)
            End If
        Next

        Return 0
    End Function

#End Region

#Region "Contrôle du type de rapport"

    ''' <summary>
    ''' Détermine le type réel d'un rapport d'après les colonnes qu'il contient.
    ''' C'est le contrôle qui FAIT FOI : contrairement au nom, le contenu ne peut pas mentir.
    ''' Retourne Indetermine si aucune des deux signatures n'est suffisamment représentée
    ''' (fichier d'une autre nature, ou format Western Union profondément remanié).
    ''' </summary>
    Public Shared Function DetecterTypeDepuisContenu(table As DataTable) As TypeRapportWU

        If table Is Nothing OrElse table.Columns.Count = 0 Then
            Return TypeRapportWU.Indetermine
        End If

        Dim scoreActivite As Integer = CompterColonnesPresentes(table, ConstantesWU.SignatureRapportActivite)
        Dim scoreReglement As Integer = CompterColonnesPresentes(table, ConstantesWU.SignatureRapportReglement)

        ' Les deux signatures n'ayant aucune colonne commune, un fichier valide ne peut pas
        ' satisfaire les deux : une égalité ne peut donc traduire qu'un fichier non reconnu.
        If scoreActivite >= ConstantesWU.MIN_COLONNES_SIGNATURE AndAlso scoreActivite > scoreReglement Then
            Return TypeRapportWU.Activite
        End If

        If scoreReglement >= ConstantesWU.MIN_COLONNES_SIGNATURE AndAlso scoreReglement > scoreActivite Then
            Return TypeRapportWU.Reglement
        End If

        Return TypeRapportWU.Indetermine
    End Function

    ''' <summary>
    ''' Vérifie qu'un rapport chargé est bien du type attendu, d'après son CONTENU.
    ''' Appelée avant tout calcul : elle empêche qu'un rapport d'activité soit exploité comme
    ''' rapport de règlement (ou l'inverse), y compris si le fichier a été renommé.
    ''' Un contenu non reconnu (Indetermine) n'est pas bloqué ici : le contrôle des colonnes
    ''' obligatoires (VerifierColonnesRapport) rendra alors un diagnostic plus précis.
    ''' </summary>
    ''' <exception cref="RapportInvalideException">Levée si le fichier est du type opposé.</exception>
    Public Shared Sub VerifierTypeRapport(table As DataTable, typeAttendu As TypeRapportWU, nomFichier As String)

        Dim typeReel As TypeRapportWU = DetecterTypeDepuisContenu(table)

        If typeReel = TypeRapportWU.Indetermine OrElse typeReel = typeAttendu Then
            Return
        End If

        Throw New RapportInvalideException(
            $"Le fichier « {nomFichier} » est en réalité un rapport {LibelleType(typeReel).ToUpperInvariant()}, " &
            $"alors qu'il a été chargé comme rapport {LibelleType(typeAttendu).ToUpperInvariant()}." &
            Environment.NewLine & Environment.NewLine &
            "Diagnostic établi sur les colonnes réellement présentes dans le fichier, et non sur son nom." &
            Environment.NewLine &
            "Rechargez le fichier attendu à l'aide du bouton correspondant.")
    End Sub

    ''' <summary>
    ''' Vérifie, dès la sélection et d'après le seul NOM du fichier, que l'utilisateur n'a pas
    ''' choisi le rapport de l'autre type. Contrôle immédiat, avant toute lecture du fichier.
    ''' Un nom non concluant est accepté sans message : le contrôle sur le contenu prendra le relais.
    ''' </summary>
    ''' <returns>False si le nom du fichier désigne le type opposé à celui attendu.</returns>
    Public Shared Function VerifierTypeAttendu(infos As InfosFichierRapport,
                                               typeAttendu As TypeRapportWU,
                                               ByRef messageErreur As String) As Boolean

        messageErreur = String.Empty

        If infos Is Nothing OrElse infos.Type = TypeRapportWU.Indetermine OrElse infos.Type = typeAttendu Then
            Return True
        End If

        messageErreur =
            $"Le fichier « {infos.NomFichier} » est un rapport {LibelleType(infos.Type).ToUpperInvariant()}." &
            Environment.NewLine & Environment.NewLine &
            $"Ce bouton attend le rapport {LibelleType(typeAttendu).ToUpperInvariant()}." &
            Environment.NewLine &
            "Sélectionnez le fichier correspondant, ou utilisez l'autre bouton de chargement."

        Return False
    End Function

#End Region

#Region "Contrôle de la période"

    ''' <summary>
    ''' Vérifie que les deux rapports couvrent exactement la même période, d'après leurs noms
    ''' de fichiers. Ne bloque que sur une contradiction CERTAINE : si la période de l'un des
    ''' deux fichiers n'est pas lisible dans son nom, la comparaison est simplement impossible
    ''' et le traitement se poursuit — les dates contenues DANS les rapports seront de toute
    ''' façon confrontées par WUReportService.ValiderCoherenceDates.
    ''' </summary>
    ''' <param name="message">
    ''' Message d'erreur si la fonction retourne False ; sinon message informatif (période
    ''' retenue, ou raison pour laquelle la comparaison n'a pas pu être faite). Peut être vide.
    ''' </param>
    ''' <returns>False uniquement si les deux périodes sont lisibles ET différentes.</returns>
    Public Shared Function VerifierMemePeriode(infosActivite As InfosFichierRapport,
                                               infosReglement As InfosFichierRapport,
                                               ByRef message As String) As Boolean

        message = String.Empty

        If infosActivite Is Nothing OrElse infosReglement Is Nothing Then
            Return True ' Un seul fichier chargé pour l'instant : rien à comparer.
        End If

        Dim periodeActivite As PeriodeRapport = infosActivite.Periode
        Dim periodeReglement As PeriodeRapport = infosReglement.Periode

        If periodeActivite Is Nothing AndAlso periodeReglement Is Nothing Then
            message = "Périodes non lisibles dans le nom des deux fichiers : " &
                      "la concordance sera vérifiée sur les dates contenues dans les rapports."
            Return True
        End If

        If periodeActivite Is Nothing Then
            message = $"Période non lisible dans le nom du rapport d'activité " &
                      $"(rapport de règlement : {periodeReglement}). " &
                      "Concordance à vérifier sur les dates contenues dans les rapports."
            Return True
        End If

        If periodeReglement Is Nothing Then
            message = $"Période non lisible dans le nom du rapport de règlement " &
                      $"(rapport d'activité : {periodeActivite}). " &
                      "Concordance à vérifier sur les dates contenues dans les rapports."
            Return True
        End If

        If Not periodeActivite.EstIdentiqueA(periodeReglement) Then
            message =
                "Les deux rapports ne couvrent pas la même période." & Environment.NewLine & Environment.NewLine &
                $"    Rapport d'activité  : {periodeActivite}" & Environment.NewLine &
                $"    Rapport de règlement : {periodeReglement}" & Environment.NewLine & Environment.NewLine &
                "Une compensation ne peut être établie qu'entre deux rapports de la même période." &
                Environment.NewLine &
                "Sélectionnez les deux fichiers correspondant à la journée à comptabiliser."
            Return False
        End If

        message = $"Période des rapports : {periodeActivite}."
        Return True
    End Function

#End Region

#Region "Utilitaires internes"

    ''' <summary>Libellé lisible d'un type de rapport, pour les messages destinés à l'utilisateur.</summary>
    Public Shared Function LibelleType(type As TypeRapportWU) As String
        Select Case type
            Case TypeRapportWU.Activite
                Return "d'activité"
            Case TypeRapportWU.Reglement
                Return "de règlement"
            Case Else
                Return "de type indéterminé"
        End Select
    End Function

    ''' <summary>
    ''' Met un texte en minuscules et lui retire ses accents, afin que la recherche de mots-clés
    ''' fonctionne quelle que soit la façon dont le fichier a été nommé ("activité", "ACTIVITE",
    ''' "Activite"). La décomposition Unicode traite tous les accents d'un coup, sans avoir à
    ''' énumérer chaque caractère accentué.
    ''' </summary>
    Private Shared Function Normaliser(texte As String) As String

        If String.IsNullOrEmpty(texte) Then Return String.Empty

        Dim decompose As String = texte.Trim().ToLowerInvariant().Normalize(NormalizationForm.FormD)
        Dim resultat As New StringBuilder(decompose.Length)

        For Each caractere As Char In decompose
            ' Les accents sont, après décomposition, des caractères de catégorie "NonSpacingMark".
            If CharUnicodeInfo.GetUnicodeCategory(caractere) <> UnicodeCategory.NonSpacingMark Then
                resultat.Append(caractere)
            End If
        Next

        Return resultat.ToString().Normalize(NormalizationForm.FormC)
    End Function

    ''' <summary>Indique si le texte normalisé contient au moins un des marqueurs recherchés.</summary>
    Private Shared Function ContientUnMarqueur(texteNormalise As String, marqueurs As String()) As Boolean
        For Each marqueur As String In marqueurs
            If texteNormalise.IndexOf(marqueur, StringComparison.Ordinal) >= 0 Then
                Return True
            End If
        Next
        Return False
    End Function

    ''' <summary>Compte, parmi une liste de colonnes, celles réellement présentes dans la table.</summary>
    Private Shared Function CompterColonnesPresentes(table As DataTable, colonnes As String()) As Integer
        Dim total As Integer = 0
        For Each colonne As String In colonnes
            If table.Columns.Contains(colonne) Then
                total += 1
            End If
        Next
        Return total
    End Function

#End Region

End Class
