Option Strict On
Option Explicit On

Imports System.Data
Imports System.Globalization

''' <summary>
''' Le bordereau de fin de journée : ce qui s'imprime, se signe et se classe.
'''
''' CE QU'IL EST, ET CE QU'IL N'EST PAS
'''
''' Ce n'est PAS une pièce comptable. La pièce porte déjà les quatre cartouches de la banque,
''' et les écritures sont couvertes par une signature. Le bordereau atteste de la FAÇON dont
''' la journée a été faite — et surtout de ce qui n'a pas pu l'être.
'''
''' SON BLOC LE PLUS IMPORTANT EST CELUI DES ACCOUNTS ÉCARTÉS
'''
''' Les points de vente non paramétrés ne figurent nulle part sur la pièce, par construction.
''' Un contrôleur qui ne voit que la pièce ne peut pas savoir qu'une agence a travaillé ce
''' jour-là sans que ses opérations soient passées. Les cinq autres blocs décrivent ce qui a
''' été fait ; celui-là décrit ce qui ne l'a pas été, et c'est la seule information qu'un
''' supérieur ne peut obtenir ailleurs.
'''
''' IL SE RECONSTITUE
'''
''' Tout vient de la base : l'en-tête de traitement, l'historique, la pièce conservée et la
''' trace du fichier core banking. Rien n'est recalculé avec le paramétrage d'aujourd'hui.
''' </summary>
Public NotInheritable Class BordereauService

    Private Sub New()
    End Sub

#Region "Ce que porte un bordereau"

    ''' <summary>Le bordereau d'une journée, prêt à être affiché ou imprimé.</summary>
    Public NotInheritable Class Contenu

        Public Property Traitement As TraitementJourneeWU
        Public Property Lignes As New List(Of LigneHistoriqueWU)()

        Public Property Identification As DataTable
        Public Property Sources As DataTable
        Public Property Volumes As DataTable
        Public Property Montants As DataTable
        Public Property Piece As DataTable
        Public Property Ecartes As DataTable
        Public Property CoreBanking As DataTable

        ''' <summary>Vrai si la journée n'a rien laissé dans l'historique.</summary>
        Public ReadOnly Property EstVide As Boolean
            Get
                Return Lignes.Count = 0
            End Get
        End Property

        ''' <summary>Titre du document, tel qu'il s'imprime en tête.</summary>
        Public ReadOnly Property Titre As String
            Get
                Return "ECOBANK TCHAD — BORDEREAU DE COMPENSATION WESTERN UNION"
            End Get
        End Property

        Public ReadOnly Property NomDeFichier As String
            Get
                Return $"Bordereau-{Traitement.DateActivite:yyyyMMdd}.pdf"
            End Get
        End Property
    End Class

#End Region

#Region "Construction"

    ''' <summary>
    ''' Rassemble le bordereau d'une journée depuis la base.
    '''
    ''' Un en-tête de traitement absent n'interrompt rien : la journée est antérieure à la mise
    ''' en service du bordereau, et tout le reste se lit quand même. Le document le dira.
    ''' </summary>
    Public Shared Function Construire(jour As Date, ByRef messageErreur As String) As Contenu

        messageErreur = String.Empty

        Dim motif As String = String.Empty

        Dim traitement As TraitementJourneeWU = TraitementRepository.Charger(jour, motif)
        If traitement Is Nothing Then
            traitement = New TraitementJourneeWU() With {.DateActivite = jour.Date}
        End If

        Dim lignes As List(Of LigneHistoriqueWU) = HistoriqueRepository.ListerPeriode(jour, jour, motif)
        If motif.Length > 0 Then
            messageErreur = motif
            Return Nothing
        End If

        Dim contenu As New Contenu() With {
            .Traitement = traitement,
            .Lignes = lignes
        }

        ' Sans en-tête enregistré, les chiffres de la pièce se relisent dans la pièce elle-même.
        If Not traitement.Enregistre Then CompleterDepuisLHistorique(contenu)

        contenu.Identification = ConstruireIdentification(contenu)
        contenu.Sources = ConstruireSources(contenu)
        contenu.Volumes = ConstruireVolumes(contenu)
        contenu.Montants = ConstruireMontants(contenu)
        contenu.Piece = ConstruirePiece(contenu)
        contenu.Ecartes = ConstruireEcartes(contenu)
        contenu.CoreBanking = ConstruireCoreBanking(contenu)

        Return contenu
    End Function

    ''' <summary>
    ''' Pour une journée antérieure à l'en-tête de traitement : retrouve ce qui peut l'être.
    '''
    ''' Les volumes et les totaux de la pièce se relisent ; le nom des rapports Western Union,
    ''' non — il n'était conservé nulle part. Le bordereau le dira au lieu de laisser des cases
    ''' vides qui se liraient comme une absence de rapport.
    ''' </summary>
    Private Shared Sub CompleterDepuisLHistorique(contenu As Contenu)

        Dim traitement As TraitementJourneeWU = contenu.Traitement
        Dim motif As String = String.Empty

        Dim piece As DataTable = PieceRepository.Charger(traitement.DateActivite, motif)

        If piece IsNot Nothing Then
            For Each ligne As DataRow In piece.Rows
                traitement.TotalDebit += Convert.ToInt64(ligne("Debit"), CultureInfo.InvariantCulture)
                traitement.TotalCredit += Convert.ToInt64(ligne("Credit"), CultureInfo.InvariantCulture)
            Next
        End If

        For Each ligne As LigneHistoriqueWU In contenu.Lignes
            traitement.NombreEnvois += ligne.NombreEnvois
            traitement.NombrePaiements += ligne.NombrePaiements
            traitement.NombreAnnulations += ligne.NombreAnnulations
        Next

        traitement.NombrePdv = contenu.Lignes.Count
        traitement.NombreSousAgents = Compter(contenu.Lignes, "SA")
        traitement.NombreAgences = Compter(contenu.Lignes, "EC")
        traitement.NombreEcartes = ComptesEcartes(contenu.Lignes).Count
        traitement.NumeroLot = CoreBankingService.NumeroDeLot(traitement.DateActivite)
    End Sub

    Private Shared Function Compter(lignes As List(Of LigneHistoriqueWU), typeCherche As String) As Integer

        Dim total As Integer = 0

        For Each ligne As LigneHistoriqueWU In lignes
            If String.Equals(ligne.TypePdv, typeCherche, StringComparison.OrdinalIgnoreCase) Then total += 1
        Next

        Return total
    End Function

    ''' <summary>
    ''' Les points de vente qui n'ont été comptabilisés NI comme sous-agent NI comme agence
    ''' propre : ceux dont la banque ne connaissait pas les comptes ce jour-là.
    ''' </summary>
    Private Shared Function ComptesEcartes(lignes As List(Of LigneHistoriqueWU)) As List(Of LigneHistoriqueWU)

        Dim ecartes As New List(Of LigneHistoriqueWU)()

        For Each ligne As LigneHistoriqueWU In lignes

            If String.Equals(ligne.TypePdv, "SA", StringComparison.OrdinalIgnoreCase) Then Continue For
            If String.Equals(ligne.TypePdv, "EC", StringComparison.OrdinalIgnoreCase) Then Continue For

            ecartes.Add(ligne)
        Next

        Return ecartes
    End Function

#End Region

#Region "Les blocs"

    Private Shared Function DeuxColonnes(premiere As String, seconde As String) As DataTable

        Dim table As New DataTable("Bloc")
        table.Columns.Add(premiere, GetType(String))
        table.Columns.Add(seconde, GetType(String))
        Return table
    End Function

    Private Shared Function ConstruireIdentification(contenu As Contenu) As DataTable

        Dim table As DataTable = DeuxColonnes("Rubrique", "Valeur")
        Dim traitement As TraitementJourneeWU = contenu.Traitement

        table.Rows.Add("Journée d'activité", Jour(traitement.DateActivite))
        table.Rows.Add("Date de valeur",
                       If(traitement.DateValeur.HasValue, Jour(traitement.DateValeur.Value), "non enregistrée"))
        table.Rows.Add("Numéro de lot", Defaut(traitement.NumeroLot))

        table.Rows.Add("Comptabilisée par", Defaut(traitement.ComptabilisePar))
        table.Rows.Add("Comptabilisée le",
                       If(traitement.DateComptabilisation.HasValue,
                          traitement.DateComptabilisation.Value.ToString("dd/MM/yyyy à HH:mm",
                                                                        CultureInfo.InvariantCulture),
                          "non enregistrée"))

        table.Rows.Add("Visa", traitement.LibelleVisa)

        If traitement.CommentaireVisa.Length > 0 Then
            table.Rows.Add("Observation du viseur", traitement.CommentaireVisa)
        End If

        Return table
    End Function

    Private Shared Function ConstruireSources(contenu As Contenu) As DataTable

        Dim table As New DataTable("Sources")
        table.Columns.Add("Rapport", GetType(String))
        table.Columns.Add("Fichier", GetType(String))
        table.Columns.Add("Empreinte SHA-256", GetType(String))

        Dim traitement As TraitementJourneeWU = contenu.Traitement

        If Not traitement.Enregistre Then
            table.Rows.Add("Rapport d'activité", "NON CONSERVÉ pour cette journée", String.Empty)
            table.Rows.Add("Rapport de règlement", "NON CONSERVÉ pour cette journée", String.Empty)
            Return table
        End If

        table.Rows.Add("Rapport d'activité", Defaut(traitement.FichierActivite),
                       Defaut(traitement.EmpreinteActivite))
        table.Rows.Add("Rapport de règlement", Defaut(traitement.FichierReglement),
                       Defaut(traitement.EmpreinteReglement))

        Return table
    End Function

    Private Shared Function ConstruireVolumes(contenu As Contenu) As DataTable

        Dim table As DataTable = DeuxColonnes("Rubrique", "Nombre")
        Dim traitement As TraitementJourneeWU = contenu.Traitement

        table.Rows.Add("Points de vente traités", Nombre(traitement.NombrePdv))
        table.Rows.Add("dont sous-agents", Nombre(traitement.NombreSousAgents))
        table.Rows.Add("dont agences propres", Nombre(traitement.NombreAgences))
        table.Rows.Add("ACCOUNTS ÉCARTÉS (non paramétrés)", Nombre(traitement.NombreEcartes))
        table.Rows.Add("Envois", Nombre(traitement.NombreEnvois))
        table.Rows.Add("Paiements", Nombre(traitement.NombrePaiements))
        table.Rows.Add("Annulations", Nombre(traitement.NombreAnnulations))

        Return table
    End Function

    Private Shared Function ConstruireMontants(contenu As Contenu) As DataTable

        Dim table As DataTable = DeuxColonnes("Rubrique", "Montant (FCFA)")

        Dim cumul As New LigneHistoriqueWU()
        For Each ligne As LigneHistoriqueWU In contenu.Lignes
            cumul.Cumuler(ligne)
        Next

        table.Rows.Add("Principal envoyé", Montant(cumul.PrincipalEnvoi))
        table.Rows.Add("Principal payé", Montant(cumul.PrincipalPaye))
        table.Rows.Add("Charges sur envoi", Montant(cumul.ChargeEnvoi))
        table.Rows.Add("Taxes perçues", Montant(cumul.TotalTaxes))
        table.Rows.Add("Commission totale", Montant(cumul.TotalCommissions))

        ' Ce que la banque garde, et ce qui part aux sous-agents : le bordereau doit permettre
        ' de retrouver l'état des commissions sans le rouvrir.
        If cumul.RepartitionConnue Then
            table.Rows.Add("dont part de la banque", Montant(cumul.TotalPartBanque))
            table.Rows.Add("dont rétrocédé aux sous-agents", Montant(cumul.TotalPartSousAgent))
        Else
            table.Rows.Add("dont part de la banque", "répartition non conservée pour cette journée")
        End If

        Return table
    End Function

    Private Shared Function ConstruirePiece(contenu As Contenu) As DataTable

        Dim table As DataTable = DeuxColonnes("Rubrique", "Valeur")
        Dim traitement As TraitementJourneeWU = contenu.Traitement

        table.Rows.Add("Total débit", Montant(traitement.TotalDebit))
        table.Rows.Add("Total crédit", Montant(traitement.TotalCredit))
        table.Rows.Add("Équilibre", If(traitement.Equilibree, "ÉQUILIBRÉE", "DÉSÉQUILIBRÉE"))

        ' L'écart d'arrondi est dans la pièce, mais noyé au milieu des lignes. Ici il est seul,
        ' avec son compte : c'est ce qu'un contrôleur regarde.
        table.Rows.Add("Écart d'arrondi", Montant(traitement.EcartArrondi))
        table.Rows.Add("Compte de l'écart", Defaut(traitement.CompteEcart))

        Return table
    End Function

    ''' <summary>
    ''' Le bloc qui justifie la signature : les points de vente qui ont travaillé ce jour-là
    ''' sans que leurs opérations soient passées en comptabilité.
    ''' </summary>
    Private Shared Function ConstruireEcartes(contenu As Contenu) As DataTable

        Dim table As New DataTable("Ecartes")
        table.Columns.Add("Account", GetType(String))
        table.Columns.Add("Designation", GetType(String))
        table.Columns.Add("Envois", GetType(Integer))
        table.Columns.Add("Paiements", GetType(Integer))
        table.Columns.Add("PrincipalEnvoi", GetType(Decimal))
        table.Columns.Add("PrincipalPaye", GetType(Decimal))
        table.Columns.Add("Commission", GetType(Decimal))

        For Each ligne As LigneHistoriqueWU In ComptesEcartes(contenu.Lignes)
            table.Rows.Add(ligne.Account, ligne.Designation, ligne.NombreEnvois, ligne.NombrePaiements,
                           ligne.PrincipalEnvoi, ligne.PrincipalPaye, ligne.TotalCommissions)
        Next

        Return table
    End Function

    Private Shared Function ConstruireCoreBanking(contenu As Contenu) As DataTable

        Dim table As DataTable = DeuxColonnes("Rubrique", "Valeur")

        Dim production As CoreBankingRepository.Production =
            CoreBankingRepository.DerniereProduction(contenu.Traitement.DateActivite)

        If production Is Nothing Then
            table.Rows.Add("Fichier core banking",
                           "Aucune production enregistrée — ce qui ne veut pas dire qu'aucun " &
                           "fichier n'est parti : les fichiers produits avant la mise en service " &
                           "de cette trace ne sont pas enregistrés.")
            Return table
        End If

        table.Rows.Add("Fichier produit", Defaut(production.NomFichier))
        table.Rows.Add("Produit le",
                       production.DateProduction.ToString("dd/MM/yyyy à HH:mm", CultureInfo.InvariantCulture))
        table.Rows.Add("Produit par", Defaut(production.ProduitPar))
        table.Rows.Add("Numéro de lot", Defaut(production.NumeroLot))
        table.Rows.Add("Date de valeur", Jour(production.DateValeur))
        table.Rows.Add("Lignes", Nombre(production.NombreLignes))
        table.Rows.Add("Total débit", Montant(production.TotalDebit))

        Return table
    End Function

#End Region

#Region "Export en PDF"

    ''' <summary>Nombre d'étapes annoncées à la barre de progression.</summary>
    Public Const ETAPES As Integer = 3

    ''' <summary>
    ''' Écrit le bordereau en PDF, cartouches de signature comprises.
    '''
    ''' En PDF et non en classeur : ce document se signe et se classe. Un tableur invite à
    ''' retoucher les chiffres, et un chiffre retouché sur le document qu'on présente n'est
    ''' plus celui de la journée.
    ''' </summary>
    Public Shared Sub ExporterEnPdf(contenu As Contenu, chemin As String,
                                    progression As ProgressionWU)

        If contenu Is Nothing Then
            Throw New InvalidOperationException("Aucun bordereau à exporter.")
        End If

        ExcelExportService.ExporterEnPdf(contenu.Titre, SousTitres(contenu), Blocs(contenu),
                                         "Bordereau", chemin, progression:=progression)
    End Sub

    Private Shared Function SousTitres(contenu As Contenu) As List(Of SousTitreExcel)

        Dim traitement As TraitementJourneeWU = contenu.Traitement

        Dim lignes As New List(Of SousTitreExcel) From {
            New SousTitreExcel(traitement.Intitule),
            New SousTitreExcel(traitement.LibelleVisa, Not traitement.EstVisee)
        }

        If traitement.NombreEcartes > 0 Then
            lignes.Add(New SousTitreExcel(
                $"{traitement.NombreEcartes} Account(s) NON COMPTABILISÉ(S) — voir le tableau " &
                "en fin de bordereau.", True))
        End If

        If Not traitement.Equilibree Then
            lignes.Add(New SousTitreExcel("PIÈCE DÉSÉQUILIBRÉE : débit et crédit ne " &
                                          "correspondent pas.", True))
        End If

        If Not traitement.Enregistre Then
            lignes.Add(New SousTitreExcel(
                "Journée comptabilisée avant la mise en service du bordereau : le nom des " &
                "rapports Western Union n'a pas été conservé.", True))
        End If

        lignes.Add(New SousTitreExcel($"Édité le {Date.Now:dd/MM/yyyy à HH:mm} par {SessionWU.Auteur}"))

        Return lignes
    End Function

    Private Shared Function Blocs(contenu As Contenu) As List(Of BlocExcel)

        Dim formats As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase) From {
            {"Envois", "# ##0"}, {"Paiements", "# ##0"},
            {"PrincipalEnvoi", "# ##0"}, {"PrincipalPaye", "# ##0"}, {"Commission", "# ##0"}
        }

        Dim entetes As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase) From {
            {"Designation", "Désignation"}, {"PrincipalEnvoi", "Principal envoyé"},
            {"PrincipalPaye", "Principal payé"}, {"Commission", "Commission générée"}
        }

        ' La variable ne s'appelle pas « blocs » : le nom d'une fonction EST sa variable de
        ' retour en Visual Basic, et une locale homonyme est refusée (BC30290).
        Dim liste As New List(Of BlocExcel) From {
            New BlocExcel("1. Identification de la journée", contenu.Identification),
            New BlocExcel("2. Rapports Western Union traités", contenu.Sources),
            New BlocExcel("3. Ce qui a été traité", contenu.Volumes),
            New BlocExcel("4. Montants de la journée", contenu.Montants),
            New BlocExcel("5. La pièce comptable", contenu.Piece)
        }

        liste.Add(BlocDesEcartes(contenu, formats, entetes))
        liste.Add(New BlocExcel("7. Fichier destiné au core banking", contenu.CoreBanking))
        liste.Add(BlocDesSignatures())

        Return liste
    End Function

    ''' <summary>
    ''' Le bloc des écartés, mis en exergue quand il n'est pas vide — et rendu explicite
    ''' quand il l'est. « Aucun » se lit ; une case vide ne se lit pas.
    ''' </summary>
    Private Shared Function BlocDesEcartes(contenu As Contenu,
                                           formats As Dictionary(Of String, String),
                                           entetes As Dictionary(Of String, String)) As BlocExcel

        If contenu.Ecartes.Rows.Count > 0 Then
            Return New BlocExcel() With {
                .Titre = "6. ACCOUNTS NON COMPTABILISÉS — points de vente non paramétrés",
                .Donnees = contenu.Ecartes,
                .Formats = formats,
                .Entetes = entetes,
                .TitreEnExergue = True
            }
        End If

        Dim aucun As DataTable = DeuxColonnes("Rubrique", "Valeur")
        aucun.Rows.Add("Accounts non comptabilisés",
                       "AUCUN — tous les points de vente de la journée étaient paramétrés.")

        Return New BlocExcel("6. Accounts non comptabilisés", aucun)
    End Function

    ''' <summary>
    ''' Les cartouches de signature. Trois lignes hautes : c'est l'espace où l'on signe.
    ''' </summary>
    Private Shared Function BlocDesSignatures() As BlocExcel

        Dim table As New DataTable("Signatures")
        table.Columns.Add("Fonction", GetType(String))
        table.Columns.Add("Nom", GetType(String))
        table.Columns.Add("Date", GetType(String))
        table.Columns.Add("Signature", GetType(String))

        table.Rows.Add("Établi par — agent de la compense", String.Empty, String.Empty, String.Empty)
        table.Rows.Add("Vérifié par — chef de service", String.Empty, String.Empty, String.Empty)
        table.Rows.Add("Approuvé par", String.Empty, String.Empty, String.Empty)

        Return New BlocExcel() With {
            .Titre = "8. Visa et signatures",
            .Donnees = table,
            .HauteurLignes = 42R
        }
    End Function

#End Region

#Region "Libellés"

    Private Shared Function Jour(valeur As Date) As String
        Return valeur.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture)
    End Function

    Private Shared Function Nombre(valeur As Integer) As String
        Return valeur.ToString("N0", CultureInfo.CurrentCulture)
    End Function

    Private Shared Function Montant(valeur As Long) As String
        Return valeur.ToString("N0", CultureInfo.CurrentCulture)
    End Function

    Private Shared Function Montant(valeur As Decimal) As String
        Return valeur.ToString("N0", CultureInfo.CurrentCulture)
    End Function

    ''' <summary>Une valeur absente se dit, elle ne se laisse pas vide.</summary>
    Private Shared Function Defaut(valeur As String) As String

        Dim texte As String = If(valeur, String.Empty).Trim()
        If texte.Length = 0 Then Return "—"
        Return texte
    End Function

#End Region

End Class
