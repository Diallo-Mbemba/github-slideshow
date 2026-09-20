Option Strict On
Option Explicit On

Imports System.Globalization
Imports System.IO
Imports System.IO.Compression
Imports System.Text

''' <summary>
''' Le FORMAT du fichier de paramétrage : ce qu'il contient, comment il s'écrit, comment il se
''' relit. Point unique de vérité — l'export et l'import passent tous deux par ici, et ne
''' peuvent donc pas diverger.
'''
''' UNE ARCHIVE DE FICHIERS TEXTE, ET NON UN CLASSEUR
'''
''' Le fichier doit pouvoir être relu sur le poste d'une installation qui commence, donc sans
''' qu'on puisse parier sur la présence d'Excel. Il doit aussi rester LISIBLE : un
''' administrateur qui ouvre l'archive doit voir ses sous-agents, pas des octets.
'''
''' D'où une archive ZIP de CSV en UTF-8, que l'on ouvre d'un double-clic, chaque fichier
''' s'ouvrant lui-même dans Excel si Excel est là. L'application, elle, n'en a pas besoin :
''' elle lit et écrit du texte.
'''
''' L'ARCHIVE SE RECHARGE TELLE QUELLE
'''
''' On charge le .zip, pas les fichiers extraits. C'est le manifeste qui dit d'où vient le lot
''' et si son contenu a été retouché ; un fichier isolé ne dirait ni l'un ni l'autre.
''' </summary>
Public NotInheritable Class ParametrageFichierWU

    Private Sub New()
    End Sub

#Region "Le format"

    ''' <summary>
    ''' Version du format. Un fichier produit par une version PLUS RÉCENTE est refusé : il
    ''' porte peut-être des colonnes que cette application ne saurait pas relire, et charger
    ''' la moitié d'un paramétrage est pire que n'en charger aucun.
    ''' </summary>
    Public Const VERSION As String = "1"

    Private Const MANIFESTE As String = "manifeste.csv"
    Private Const LISEZ_MOI As String = "LISEZ-MOI.txt"

    Private Const COMPTES As String = "comptes-systeme.csv"
    Private Const GROUPES As String = "groupes-statistiques.csv"
    Private Const SOUS_AGENTS As String = "sous-agents.csv"
    Private Const AGENCES As String = "agences-propres.csv"
    Private Const UTILISATEURS As String = "utilisateurs-pour-information.csv"

    ''' <summary>Nom proposé pour l'archive : la base et l'instant, pour qu'on s'y retrouve.</summary>
    Public Shared Function NomDeFichier(base As String) As String

        Dim nomBase As String = If(base, String.Empty).Trim()
        If nomBase.Length = 0 Then nomBase = "Wincompense"

        Return $"Parametrage-{nomBase}-{Date.Now:yyyyMMdd-HHmm}.zip"
    End Function

#End Region

#Region "Écriture"

    ''' <summary>
    ''' Écrit l'archive. Chaque table devient un fichier texte ; le manifeste les résume et
    ''' porte l'empreinte de leur contenu.
    ''' </summary>
    Public Shared Function Ecrire(lot As LotParametrageWU, chemin As String,
                                  progression As ProgressionWU,
                                  ByRef messageErreur As String) As Boolean

        messageErreur = String.Empty

        If lot Is Nothing Then
            messageErreur = "Aucun paramétrage à écrire."
            Return False
        End If

        Try
            ' Un fichier laissé par une tentative précédente ferait une archive incohérente :
            ' ZipArchive en mode Create écrit par-dessus, mais sans effacer ce qui dépasse.
            If File.Exists(chemin) Then File.Delete(chemin)

            Dim contenus As New Dictionary(Of String, String)(StringComparer.Ordinal) From {
                {COMPTES, TexteDesComptes(lot)},
                {GROUPES, TexteDesGroupes(lot)},
                {SOUS_AGENTS, TexteDesSousAgents(lot)},
                {AGENCES, TexteDesAgences(lot)}
            }

            Annoncer(progression, "Écriture de l'archive")

            Using flux As FileStream = File.Create(chemin)
                Using archive As New ZipArchive(flux, ZipArchiveMode.Create)

                    EcrireEntree(archive, MANIFESTE, TexteDuManifeste(lot, contenus))
                    EcrireEntree(archive, LISEZ_MOI, TexteDuLisezMoi(lot))

                    For Each entree As KeyValuePair(Of String, String) In contenus
                        EcrireEntree(archive, entree.Key, entree.Value)
                    Next

                    EcrireEntree(archive, UTILISATEURS, TexteDesUtilisateurs(lot))
                End Using
            End Using

        Catch ex As IOException
            messageErreur = $"Écriture du fichier impossible : {ex.Message}"
            Return False
        Catch ex As UnauthorizedAccessException
            messageErreur = $"Écriture refusée par Windows : {ex.Message}"
            Return False
        End Try

        Return True
    End Function

    Private Shared Sub EcrireEntree(archive As ZipArchive, nom As String, contenu As String)

        Dim entree As ZipArchiveEntry = archive.CreateEntry(nom, CompressionLevel.Optimal)

        Using flux As Stream = entree.Open()
            Using plume As New StreamWriter(flux, FichierCsvWU.Encodage)
                plume.Write(contenu)
            End Using
        End Using
    End Sub

    Private Shared Function TexteDuManifeste(lot As LotParametrageWU,
                                             contenus As Dictionary(Of String, String)) As String

        Dim assemblage As New StringBuilder()

        assemblage.AppendLine(FichierCsvWU.Ligne("Parametre", "Valeur"))
        assemblage.AppendLine(FichierCsvWU.Ligne("Version", VERSION))
        assemblage.AppendLine(FichierCsvWU.Ligne("Base", lot.Base))
        assemblage.AppendLine(FichierCsvWU.Ligne("Serveur", lot.Serveur))
        assemblage.AppendLine(FichierCsvWU.Ligne("DateExport",
                                                 lot.DateExport.ToString("yyyy-MM-dd HH:mm:ss",
                                                                         CultureInfo.InvariantCulture)))
        assemblage.AppendLine(FichierCsvWU.Ligne("ExportePar", lot.ExportePar))
        assemblage.AppendLine(FichierCsvWU.Ligne("Empreinte", EmpreinteDuContenu(contenus)))

        Return assemblage.ToString()
    End Function

    ''' <summary>
    ''' Empreinte des QUATRE fichiers de données, dans un ordre fixe.
    '''
    ''' Le manifeste et le LISEZ-MOI en sont exclus : le premier la porte, le second ne dit
    ''' rien de ce qui sera chargé. La liste des utilisateurs aussi : elle n'est jamais
    ''' rechargée, et la retoucher ne change rien à ce qui entrera en base.
    ''' </summary>
    Private Shared Function EmpreinteDuContenu(contenus As Dictionary(Of String, String)) As String

        Dim assemblage As New StringBuilder()

        For Each nom As String In New String() {COMPTES, GROUPES, SOUS_AGENTS, AGENCES}
            assemblage.Append(nom)
            assemblage.Append(ControlChars.Lf)
            assemblage.Append(If(contenus.ContainsKey(nom), contenus(nom), String.Empty))
            assemblage.Append(ControlChars.Lf)
        Next

        Return FichierCsvWU.Empreinte(assemblage.ToString())
    End Function

    Private Shared Function TexteDesComptes(lot As LotParametrageWU) As String

        Dim assemblage As New StringBuilder()
        assemblage.AppendLine(FichierCsvWU.Ligne("Parametre", "Compte", "Signification"))

        Dim comptes As ComptesSystemeWU = lot.Comptes
        If comptes Is Nothing Then Return assemblage.ToString()

        ' Une ligne par compte, et non une colonne par compte : ajouter un compte demain
        ' n'invalidera pas les fichiers d'aujourd'hui, qui porteront simplement une ligne
        ' de moins.
        For Each ligne As String() In ParametrageColonnes.Paires(comptes)
            assemblage.AppendLine(FichierCsvWU.Ligne(ligne(0), ligne(1), ligne(2)))
        Next

        Return assemblage.ToString()
    End Function

    Private Shared Function TexteDesGroupes(lot As LotParametrageWU) As String

        Dim assemblage As New StringBuilder()
        assemblage.AppendLine(FichierCsvWU.Ligne("Groupe", "CompteActivite", "CompteCommission", "Taux"))

        For Each groupe As GroupeStatistiqueWU In lot.Groupes
            If groupe Is Nothing Then Continue For
            assemblage.AppendLine(FichierCsvWU.Ligne(groupe.Nom, groupe.CompteActivite,
                                                     groupe.CompteCommission,
                                                     FichierCsvWU.Nombre(groupe.Taux)))
        Next

        Return assemblage.ToString()
    End Function

    Private Shared Function TexteDesSousAgents(lot As LotParametrageWU) As String

        Dim assemblage As New StringBuilder()
        assemblage.AppendLine(FichierCsvWU.Ligne("Code_Pdv", "Designationagence", "GroupeStatistique",
                                                 "Taux", "CompteCompense", "CompteCommission", "codeagence"))

        For Each pdv As PointDeVenteSA In lot.SousAgents
            If pdv Is Nothing Then Continue For
            assemblage.AppendLine(FichierCsvWU.Ligne(pdv.CodePdv, pdv.Designation, pdv.GroupeStatistique,
                                                     FichierCsvWU.Nombre(pdv.Taux), pdv.CompteCompense,
                                                     pdv.CompteCommission, pdv.CodeAgence))
        Next

        Return assemblage.ToString()
    End Function

    Private Shared Function TexteDesAgences(lot As LotParametrageWU) As String

        Dim assemblage As New StringBuilder()
        assemblage.AppendLine(FichierCsvWU.Ligne("Codesite", "Designationagence", "CodeAgenc-Voyager"))

        For Each pdv As PointDeVenteEC In lot.Agences
            If pdv Is Nothing Then Continue For
            assemblage.AppendLine(FichierCsvWU.Ligne(pdv.CodeSite, pdv.Designation, pdv.CodeAgenceVoyager))
        Next

        Return assemblage.ToString()
    End Function

    ''' <summary>
    ''' La liste des utilisateurs, SANS empreinte ni sel. Elle dit qui recréer, pas comment
    ''' se connecter à leur place.
    ''' </summary>
    Private Shared Function TexteDesUtilisateurs(lot As LotParametrageWU) As String

        Dim assemblage As New StringBuilder()
        assemblage.AppendLine(FichierCsvWU.Ligne("Identifiant", "NomComplet", "Role", "Fonction", "Actif"))

        For Each utilisateur As UtilisateurWU In lot.Utilisateurs
            If utilisateur Is Nothing Then Continue For
            assemblage.AppendLine(FichierCsvWU.Ligne(utilisateur.Identifiant, utilisateur.NomComplet,
                                                     utilisateur.LibelleRole, utilisateur.LibelleFonction,
                                                     FichierCsvWU.Booleen(utilisateur.Actif)))
        Next

        Return assemblage.ToString()
    End Function

    Private Shared Function TexteDuLisezMoi(lot As LotParametrageWU) As String

        Dim assemblage As New StringBuilder()

        assemblage.AppendLine("PARAMÉTRAGE WINCOMPENSE TCHAD")
        assemblage.AppendLine("=============================")
        assemblage.AppendLine()
        assemblage.AppendLine(lot.Provenance)
        assemblage.AppendLine("Contenu : " & lot.Intitule & ".")
        assemblage.AppendLine()
        assemblage.AppendLine("CE QUE CE FICHIER EST")
        assemblage.AppendLine()
        assemblage.AppendLine("De quoi remettre en marche une installation NEUVE sans tout resaisir :")
        assemblage.AppendLine("les comptes comptables, les groupes statistiques, les sous-agents et les")
        assemblage.AppendLine("agences propres.")
        assemblage.AppendLine()
        assemblage.AppendLine("CE QUE CE FICHIER N'EST PAS")
        assemblage.AppendLine()
        assemblage.AppendLine("Une sauvegarde de la base. Il ne contient ni l'historique, ni les pièces")
        assemblage.AppendLine("comptables, ni les demandes du double regard. La sauvegarde d'une base SQL")
        assemblage.AppendLine("Server se fait sur le serveur, par l'équipe qui le tient : aucune application")
        assemblage.AppendLine("de poste ne peut la remplacer.")
        assemblage.AppendLine()
        assemblage.AppendLine("LES UTILISATEURS NE SE RECHARGENT PAS")
        assemblage.AppendLine()
        assemblage.AppendLine("Leur liste figure ici pour mémoire, pour savoir QUI recréer. Aucun mot de")
        assemblage.AppendLine("passe n'y figure, pas même sous forme d'empreinte : un fichier qui circule ne")
        assemblage.AppendLine("porte pas les identifiants d'une banque.")
        assemblage.AppendLine()
        assemblage.AppendLine("POUR LE RECHARGER")
        assemblage.AppendLine()
        assemblage.AppendLine("Wincompense, menu Paramétrage > Paramétrage : fichier de secours, onglet")
        assemblage.AppendLine("« Charger ». Choisissez CETTE ARCHIVE .zip, et non un fichier extrait.")
        assemblage.AppendLine()
        assemblage.AppendLine("Le chargement ne remplace JAMAIS une ligne déjà présente : il ne crée que ce")
        assemblage.AppendLine("qui manque. Seuls les comptes comptables font exception, et l'application le")
        assemblage.AppendLine("demande alors explicitement, en montrant l'ancienne et la nouvelle valeur.")

        Return assemblage.ToString()
    End Function

#End Region

#Region "Lecture"

    ''' <summary>
    ''' Relit une archive. Retourne Nothing et un message si elle n'est pas exploitable.
    ''' </summary>
    Public Shared Function Lire(chemin As String, ByRef messageErreur As String) As LotParametrageWU

        messageErreur = String.Empty

        Dim entrees As Dictionary(Of String, String)

        Try
            entrees = LireLesEntrees(chemin)

        Catch ex As InvalidDataException
            messageErreur = "Ce fichier n'est pas une archive lisible." & Environment.NewLine &
                            "Choisissez l'archive .zip produite par l'export, et non un fichier extrait."
            Return Nothing
        Catch ex As IOException
            messageErreur = $"Lecture du fichier impossible : {ex.Message}"
            Return Nothing
        Catch ex As UnauthorizedAccessException
            messageErreur = $"Lecture refusée par Windows : {ex.Message}"
            Return Nothing
        End Try

        If Not entrees.ContainsKey(MANIFESTE) Then
            messageErreur = "Cette archive ne porte pas de manifeste : elle n'a pas été produite " &
                            "par Wincompense, ou elle a été reconstituée à la main." & Environment.NewLine &
                            "Le chargement s'arrête : on ne charge pas un paramétrage dont on ne sait " &
                            "pas d'où il vient."
            Return Nothing
        End If

        Dim lot As New LotParametrageWU()
        LireLeManifeste(entrees(MANIFESTE), lot)

        If Not VersionAcceptee(lot.Version) Then
            messageErreur = $"Ce fichier est au format version « {lot.Version} », et cette " &
                            $"application lit la version {VERSION}." & Environment.NewLine &
                            "Il a été produit par une version plus récente de Wincompense. " &
                            "Mettez l'application à jour avant de le charger : charger la moitié " &
                            "d'un paramétrage serait pire que n'en charger aucun."
            Return Nothing
        End If

        LireLesComptes(Contenu(entrees, COMPTES), lot)
        LireLesGroupes(Contenu(entrees, GROUPES), lot)
        LireLesSousAgents(Contenu(entrees, SOUS_AGENTS), lot)
        LireLesAgences(Contenu(entrees, AGENCES), lot)

        lot.Modifie = Not EmpreinteConforme(entrees, lot)

        Return lot
    End Function

    Private Shared Function LireLesEntrees(chemin As String) As Dictionary(Of String, String)

        Dim entrees As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

        Using flux As FileStream = File.OpenRead(chemin)
            Using archive As New ZipArchive(flux, ZipArchiveMode.Read)

                For Each entree As ZipArchiveEntry In archive.Entries

                    ' Le nom seul : une archive rezippée peut avoir glissé ses fichiers dans
                    ' un sous-dossier portant le nom de l'archive.
                    Dim nom As String = Path.GetFileName(entree.FullName)
                    If nom.Length = 0 Then Continue For

                    Using lecteur As New StreamReader(entree.Open(), Encoding.UTF8, True)
                        entrees(nom) = lecteur.ReadToEnd()
                    End Using
                Next
            End Using
        End Using

        Return entrees
    End Function

    Private Shared Function Contenu(entrees As Dictionary(Of String, String), nom As String) As String
        If Not entrees.ContainsKey(nom) Then Return String.Empty
        Return entrees(nom)
    End Function

    ''' <summary>
    ''' Une version égale passe ; une version antérieure passe aussi, l'application sachant
    ''' lire ce qu'elle a écrit hier. Une version postérieure, ou illisible, ne passe pas.
    ''' </summary>
    Private Shared Function VersionAcceptee(version As String) As Boolean

        Dim lue As Integer
        Dim connue As Integer

        If Not Integer.TryParse(If(version, String.Empty).Trim(), NumberStyles.Integer,
                                CultureInfo.InvariantCulture, lue) Then Return False

        If Not Integer.TryParse(VERSION, NumberStyles.Integer, CultureInfo.InvariantCulture, connue) Then
            Return False
        End If

        Return lue <= connue
    End Function

    Private Shared Sub LireLeManifeste(texte As String, lot As LotParametrageWU)

        For Each ligne As List(Of String) In FichierCsvWU.Lire(texte)

            Dim cle As String = FichierCsvWU.Valeur(ligne, 0)
            Dim valeur As String = FichierCsvWU.Valeur(ligne, 1)

            Select Case cle
                Case "Version" : lot.Version = valeur
                Case "Base" : lot.Base = valeur
                Case "Serveur" : lot.Serveur = valeur
                Case "ExportePar" : lot.ExportePar = valeur

                Case "DateExport"
                    Dim quand As Date
                    If Date.TryParseExact(valeur, "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture,
                                          DateTimeStyles.None, quand) Then lot.DateExport = quand
            End Select
        Next
    End Sub

    Private Shared Function EmpreinteConforme(entrees As Dictionary(Of String, String),
                                              lot As LotParametrageWU) As Boolean

        Dim inscrite As String = String.Empty

        For Each ligne As List(Of String) In FichierCsvWU.Lire(Contenu(entrees, MANIFESTE))
            If FichierCsvWU.Valeur(ligne, 0) = "Empreinte" Then inscrite = FichierCsvWU.Valeur(ligne, 1)
        Next

        If inscrite.Length = 0 Then Return False

        Dim contenus As New Dictionary(Of String, String)(StringComparer.Ordinal) From {
            {COMPTES, Contenu(entrees, COMPTES)},
            {GROUPES, Contenu(entrees, GROUPES)},
            {SOUS_AGENTS, Contenu(entrees, SOUS_AGENTS)},
            {AGENCES, Contenu(entrees, AGENCES)}
        }

        Return String.Equals(inscrite, EmpreinteDuContenu(contenus), StringComparison.OrdinalIgnoreCase)
    End Function

    Private Shared Sub LireLesComptes(texte As String, lot As LotParametrageWU)

        Dim paires As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
        Dim premiere As Boolean = True

        For Each ligne As List(Of String) In FichierCsvWU.Lire(texte)

            If premiere Then
                premiere = False
                Continue For
            End If

            Dim cle As String = FichierCsvWU.Valeur(ligne, 0)
            If cle.Length = 0 Then Continue For

            paires(cle) = FichierCsvWU.Valeur(ligne, 1)
        Next

        If paires.Count = 0 Then Return
        lot.Comptes = ParametrageColonnes.Comptes(paires)
    End Sub

    Private Shared Sub LireLesGroupes(texte As String, lot As LotParametrageWU)

        Dim premiere As Boolean = True

        For Each ligne As List(Of String) In FichierCsvWU.Lire(texte)

            If premiere Then
                premiere = False
                Continue For
            End If

            Dim nom As String = FichierCsvWU.Valeur(ligne, 0)
            If nom.Length = 0 Then Continue For

            Dim taux As Decimal
            FichierCsvWU.EssayerNombre(FichierCsvWU.Valeur(ligne, 3), taux)

            lot.Groupes.Add(New GroupeStatistiqueWU() With {
                .Nom = nom,
                .CompteActivite = FichierCsvWU.Valeur(ligne, 1),
                .CompteCommission = FichierCsvWU.Valeur(ligne, 2),
                .Taux = taux
            })
        Next
    End Sub

    Private Shared Sub LireLesSousAgents(texte As String, lot As LotParametrageWU)

        Dim premiere As Boolean = True

        For Each ligne As List(Of String) In FichierCsvWU.Lire(texte)

            If premiere Then
                premiere = False
                Continue For
            End If

            Dim code As String = FichierCsvWU.Valeur(ligne, 0)
            If code.Length = 0 Then Continue For

            Dim taux As Decimal
            FichierCsvWU.EssayerNombre(FichierCsvWU.Valeur(ligne, 3), taux)

            lot.SousAgents.Add(New PointDeVenteSA() With {
                .CodePdv = code,
                .Designation = FichierCsvWU.Valeur(ligne, 1),
                .GroupeStatistique = FichierCsvWU.Valeur(ligne, 2),
                .Taux = taux,
                .CompteCompense = FichierCsvWU.Valeur(ligne, 4),
                .CompteCommission = FichierCsvWU.Valeur(ligne, 5),
                .CodeAgence = FichierCsvWU.Valeur(ligne, 6)
            })
        Next
    End Sub

    Private Shared Sub LireLesAgences(texte As String, lot As LotParametrageWU)

        Dim premiere As Boolean = True

        For Each ligne As List(Of String) In FichierCsvWU.Lire(texte)

            If premiere Then
                premiere = False
                Continue For
            End If

            Dim code As String = FichierCsvWU.Valeur(ligne, 0)
            If code.Length = 0 Then Continue For

            lot.Agences.Add(New PointDeVenteEC() With {
                .CodeSite = code,
                .Designation = FichierCsvWU.Valeur(ligne, 1),
                .CodeAgenceVoyager = FichierCsvWU.Valeur(ligne, 2)
            })
        Next
    End Sub

#End Region

#Region "Avancement"

    Private Shared Sub Annoncer(progression As ProgressionWU, libelle As String)
        If progression Is Nothing Then Return
        progression.Avancer(libelle)
    End Sub

#End Region

End Class
