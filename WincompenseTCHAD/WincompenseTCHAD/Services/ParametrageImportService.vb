Option Strict On
Option Explicit On

Imports System.Data.SqlClient
Imports System.Linq

''' <summary>
''' Chargement d'un paramétrage dans une base, pour une installation qui commence.
'''
''' LA RÈGLE QUI TIENT TOUT LE RESTE : ON NE REMPLACE JAMAIS
'''
''' Une clé déjà présente est laissée telle quelle et comptée comme « déjà présente ». Le
''' chargement ne peut donc pas détruire un référentiel : au pire, il ne fait rien.
'''
''' C'est aussi ce qui le dispense du double regard. Le contrôle à deux personnes protège les
''' MODIFICATIONS d'un paramétrage en service ; ici, il n'y a rien à protéger — là où quelque
''' chose existe, l'import s'abstient. Toute modification ultérieure repasse, elle, par la file
''' des demandes, comme aujourd'hui.
'''
''' LA SEULE EXCEPTION : LES COMPTES COMPTABLES
'''
''' SystemeWU n'est pas une liste d'objets mais UNE ligne de réglages, que le script
''' d'installation crée d'emblée avec des valeurs par défaut. S'en tenir à « ne jamais
''' remplacer » reviendrait à ne jamais charger les comptes de la banque — exactement ce qu'on
''' est venu chercher. Ils sont donc remplaçables, mais jamais en silence : l'écran montre
''' l'ancienne et la nouvelle valeur, ligne à ligne, et l'administrateur coche.
'''
''' L'ORDRE DES ÉCRITURES N'EST PAS INDIFFÉRENT
'''
''' Les groupes d'abord — les sous-agents s'y réfèrent —, puis les agences, puis les
''' sous-agents, le tout dans UNE transaction. Les comptes comptables sont écrits en DERNIER,
''' hors de cette transaction : si le référentiel échoue, les comptes n'auront pas bougé ; si
''' les comptes échouent après, le référentiel est en place et les neuf comptes se retapent
''' sur un écran, là où trois cents sous-agents ne se retapent pas.
''' </summary>
Public NotInheritable Class ParametrageImportService

    Private Sub New()
    End Sub

#Region "Ce que devient chaque ligne du fichier"

    Public Enum EtatLigneWU
        ACreer = 0
        DejaPresente = 1
        Refusee = 2
    End Enum

    ''' <summary>Une ligne du fichier, et ce que le chargement en fera.</summary>
    Public NotInheritable Class LigneRapport

        Public Property Objet As String = String.Empty
        Public Property Cle As String = String.Empty
        Public Property Designation As String = String.Empty
        Public Property Etat As EtatLigneWU = EtatLigneWU.ACreer

        ''' <summary>Pourquoi elle est refusée — ou, sur une ligne à créer, ce qui mérite un œil.</summary>
        Public Property Motif As String = String.Empty

        Public ReadOnly Property LibelleEtat As String
            Get
                Select Case Etat
                    Case EtatLigneWU.ACreer : Return "à créer"
                    Case EtatLigneWU.DejaPresente : Return "déjà présente"
                    Case EtatLigneWU.Refusee : Return "REFUSÉE"
                    Case Else : Return String.Empty
                End Select
            End Get
        End Property
    End Class

    ''' <summary>Ce que le chargement fera, avant qu'il le fasse.</summary>
    Public NotInheritable Class Rapport

        Public Property Lignes As New List(Of LigneRapport)()
        Public Property Differences As New List(Of ParametrageColonnes.Difference)()

        Public ReadOnly Property ACreer As Integer
            Get
                Return Compter(EtatLigneWU.ACreer)
            End Get
        End Property

        Public ReadOnly Property DejaPresentes As Integer
            Get
                Return Compter(EtatLigneWU.DejaPresente)
            End Get
        End Property

        Public ReadOnly Property Refusees As Integer
            Get
                Return Compter(EtatLigneWU.Refusee)
            End Get
        End Property

        Private Function Compter(etatCherche As EtatLigneWU) As Integer

            Dim total As Integer = 0
            For Each ligne As LigneRapport In Lignes
                If ligne.Etat = etatCherche Then total += 1
            Next
            Return total
        End Function

        ''' <summary>Les comptes du fichier diffèrent-ils de ceux de la base ?</summary>
        Public ReadOnly Property ComptesDifferents As Boolean
            Get
                For Each ecart As ParametrageColonnes.Difference In Differences
                    If ecart.Change Then Return True
                Next
                Return False
            End Get
        End Property

        ''' <summary>Vrai s'il n'y a rien à faire : tout est déjà là, ou tout est refusé.</summary>
        Public ReadOnly Property RienAFaire As Boolean
            Get
                Return ACreer = 0 AndAlso Not ComptesDifferents
            End Get
        End Property

        Public ReadOnly Property Intitule As String
            Get
                Return $"{ACreer} à créer, {DejaPresentes} déjà présente(s), {Refusees} refusée(s)."
            End Get
        End Property
    End Class

#End Region

#Region "Analyse"

    ''' <summary>
    ''' Confronte le fichier à la base, SANS RIEN ÉCRIRE, et dit ce qui se passerait.
    '''
    ''' L'analyse est refaite juste avant d'appliquer : entre les deux, quelqu'un d'autre a pu
    ''' créer un sous-agent. C'est le seul moyen de ne pas se heurter à une clé dupliquée au
    ''' milieu de la transaction.
    ''' </summary>
    Public Shared Function Analyser(lot As LotParametrageWU, ByRef messageErreur As String) As Rapport

        messageErreur = String.Empty

        Dim rapport As New Rapport()
        If lot Is Nothing Then
            messageErreur = "Aucun paramétrage à analyser."
            Return rapport
        End If

        Dim motif As String = String.Empty

        Dim groupesBase As List(Of GroupeStatistiqueWU) = PdvRepository.ListerGroupes(String.Empty, motif)
        If motif.Length > 0 Then
            messageErreur = "Lecture des groupes impossible : " & motif
            Return rapport
        End If

        Dim sousAgentsBase As List(Of PointDeVenteSA) = PdvRepository.ListerSousAgents(String.Empty, motif)
        If motif.Length > 0 Then
            messageErreur = "Lecture des sous-agents impossible : " & motif
            Return rapport
        End If

        Dim agencesBase As List(Of PointDeVenteEC) = PdvRepository.ListerAgences(String.Empty, motif)
        If motif.Length > 0 Then
            messageErreur = "Lecture des agences impossible : " & motif
            Return rapport
        End If

        Dim nomsGroupes As HashSet(Of String) = Ensemble(groupesBase.Select(Function(g) g.Nom))
        Dim comptesPris As HashSet(Of String) = Ensemble(
            groupesBase.Select(Function(g) g.CompteActivite).
                        Concat(groupesBase.Select(Function(g) g.CompteCommission)))

        Dim codesSA As HashSet(Of String) = Ensemble(sousAgentsBase.Select(Function(p) p.CodePdv))
        Dim codesEC As HashSet(Of String) = Ensemble(agencesBase.Select(Function(p) p.CodeSite))

        AnalyserLesGroupes(lot, rapport, nomsGroupes, comptesPris)
        AnalyserLesAgences(lot, rapport, codesEC, codesSA)
        AnalyserLesSousAgents(lot, rapport, codesSA, codesEC, nomsGroupes)

        AnalyserLesComptes(lot, rapport)

        Return rapport
    End Function

    Private Shared Sub AnalyserLesGroupes(lot As LotParametrageWU, rapport As Rapport,
                                          nomsGroupes As HashSet(Of String),
                                          comptesPris As HashSet(Of String))

        For Each groupe As GroupeStatistiqueWU In lot.Groupes

            Dim ligne As New LigneRapport() With {
                .Objet = "Groupe",
                .Cle = groupe.Nom,
                .Designation = $"{groupe.CompteActivite} / {groupe.CompteCommission}"
            }

            rapport.Lignes.Add(ligne)

            Dim anomalies As List(Of String) = groupe.Anomalies()
            If anomalies.Count > 0 Then
                Refuser(ligne, anomalies(0))
                Continue For
            End If

            If nomsGroupes.Contains(groupe.Nom) Then
                ligne.Etat = EtatLigneWU.DejaPresente
                Continue For
            End If

            ' Deux index uniques garantissent qu'un compte d'activité ou de commission
            ' n'appartient qu'à un seul groupe. Les heurter en pleine transaction ferait
            ' échouer tout le chargement : autant le dire maintenant.
            If comptesPris.Contains(groupe.CompteActivite) Then
                Refuser(ligne, $"le compte d'activité {groupe.CompteActivite} appartient déjà à un autre groupe")
                Continue For
            End If

            If comptesPris.Contains(groupe.CompteCommission) Then
                Refuser(ligne, $"le compte de commission {groupe.CompteCommission} appartient déjà à un autre groupe")
                Continue For
            End If

            nomsGroupes.Add(groupe.Nom)
            comptesPris.Add(groupe.CompteActivite)
            comptesPris.Add(groupe.CompteCommission)
        Next
    End Sub

    Private Shared Sub AnalyserLesAgences(lot As LotParametrageWU, rapport As Rapport,
                                          codesEC As HashSet(Of String),
                                          codesSA As HashSet(Of String))

        For Each agence As PointDeVenteEC In lot.Agences

            Dim ligne As New LigneRapport() With {
                .Objet = "Agence propre",
                .Cle = agence.CodeSite,
                .Designation = agence.Designation
            }

            rapport.Lignes.Add(ligne)

            Dim anomalies As List(Of String) = agence.Anomalies()
            If anomalies.Count > 0 Then
                Refuser(ligne, anomalies(0))
                Continue For
            End If

            If codesEC.Contains(agence.CodeSite) Then
                ligne.Etat = EtatLigneWU.DejaPresente
                Continue For
            End If

            ' Un Account est sous-agent OU agence propre, jamais les deux : la comptabilisation
            ' ne saurait pas quel taux appliquer.
            If codesSA.Contains(agence.CodeSite) Then
                Refuser(ligne, "cet Account est déjà enregistré comme sous-agent")
                Continue For
            End If

            codesEC.Add(agence.CodeSite)
        Next
    End Sub

    Private Shared Sub AnalyserLesSousAgents(lot As LotParametrageWU, rapport As Rapport,
                                             codesSA As HashSet(Of String),
                                             codesEC As HashSet(Of String),
                                             nomsGroupes As HashSet(Of String))

        For Each pdv As PointDeVenteSA In lot.SousAgents

            Dim ligne As New LigneRapport() With {
                .Objet = "Sous-agent",
                .Cle = pdv.CodePdv,
                .Designation = pdv.Designation
            }

            rapport.Lignes.Add(ligne)

            Dim anomalies As List(Of String) = pdv.Anomalies()
            If anomalies.Count > 0 Then
                Refuser(ligne, anomalies(0))
                Continue For
            End If

            If codesSA.Contains(pdv.CodePdv) Then
                ligne.Etat = EtatLigneWU.DejaPresente
                Continue For
            End If

            If codesEC.Contains(pdv.CodePdv) Then
                Refuser(ligne, "cet Account est déjà enregistré comme agence propre")
                Continue For
            End If

            ' Un groupe absent n'empêche pas la création — aucune clé étrangère ne l'impose
            ' aujourd'hui, et le sous-agent porte de toute façon ses propres comptes et son
            ' propre taux. Mais il le signale : le rapport d'activité le rangera à part.
            If pdv.GroupeStatistique.Length > 0 AndAlso Not nomsGroupes.Contains(pdv.GroupeStatistique) Then
                ligne.Motif = $"son groupe « {pdv.GroupeStatistique} » n'existe pas — à créer ensuite"
            ElseIf pdv.GroupeStatistique.Length = 0 Then
                ligne.Motif = "aucun groupe statistique"
            End If

            codesSA.Add(pdv.CodePdv)
        Next
    End Sub

    Private Shared Sub AnalyserLesComptes(lot As LotParametrageWU, rapport As Rapport)

        If lot.Comptes Is Nothing Then Return

        Dim motif As String = String.Empty
        Dim actuels As ComptesSystemeWU = WURepository.ChargerComptesSysteme(motif)

        rapport.Differences = ParametrageColonnes.Comparer(actuels, lot.Comptes)
    End Sub

    Private Shared Sub Refuser(ligne As LigneRapport, motif As String)
        ligne.Etat = EtatLigneWU.Refusee
        ligne.Motif = motif
    End Sub

    ''' <summary>
    ''' Les valeurs non vides, comparables sans tenir compte de la casse.
    '''
    ''' La variable locale ne s'appelle pas « ensemble » : en Visual Basic, le nom d'une
    ''' fonction EST sa variable de retour, et une locale homonyme est refusée par le
    ''' compilateur (BC30290) — la casse n'y change rien.
    ''' </summary>
    Private Shared Function Ensemble(valeurs As IEnumerable(Of String)) As HashSet(Of String)

        Dim resultat As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        For Each valeur As String In valeurs
            Dim texte As String = If(valeur, String.Empty).Trim()
            If texte.Length > 0 Then resultat.Add(texte)
        Next

        Return resultat
    End Function

#End Region

#Region "Application"

    ''' <summary>
    ''' Écrit dans la base ce que le rapport annonce, et rien d'autre.
    '''
    ''' Le référentiel tient dans UNE transaction : un chargement à moitié fait laisserait des
    ''' sous-agents rattachés à des groupes absents, et personne ne saurait dire où il s'est
    ''' arrêté. Les comptes comptables viennent après, pour la raison expliquée en tête de
    ''' classe.
    ''' </summary>
    ''' <param name="remplacerComptes">
    ''' Décision explicite de l'administrateur, prise en voyant l'ancienne et la nouvelle
    ''' valeur de chaque compte.
    ''' </param>
    Public Shared Function Appliquer(lot As LotParametrageWU, rapport As Rapport,
                                     remplacerComptes As Boolean, progression As ProgressionWU,
                                     ByRef nombreCrees As Integer,
                                     ByRef messageErreur As String) As Boolean

        messageErreur = String.Empty
        nombreCrees = 0

        If lot Is Nothing OrElse rapport Is Nothing Then
            messageErreur = "Aucun paramétrage à charger."
            Return False
        End If

        Dim aCreer As HashSet(Of String) = ClesACreer(rapport)

        Try
            Using connexion As SqlConnection = WURepository.CreerConnexion()
                connexion.Open()

                Using transaction As SqlTransaction = connexion.BeginTransaction()
                    Try
                        nombreCrees += EcrireLesGroupes(lot, aCreer, connexion, transaction, progression)
                        nombreCrees += EcrireLesAgences(lot, aCreer, connexion, transaction, progression)
                        nombreCrees += EcrireLesSousAgents(lot, aCreer, connexion, transaction, progression)

                        transaction.Commit()

                    Catch
                        Try
                            transaction.Rollback()
                        Catch
                            ' La connexion est peut-être déjà tombée : l'erreur d'origine prime.
                        End Try
                        Throw
                    End Try
                End Using
            End Using

        Catch ex As ApplicationException
            nombreCrees = 0
            messageErreur = "Chargement interrompu, RIEN n'a été écrit :" & Environment.NewLine &
                            ex.Message
            Return False

        Catch ex As SqlException
            nombreCrees = 0
            messageErreur = "Chargement interrompu, RIEN n'a été écrit :" & Environment.NewLine &
                            ex.Message
            Return False

        Catch ex As InvalidOperationException
            nombreCrees = 0
            messageErreur = $"Connexion SQL Server indisponible : {ex.Message}"
            Return False
        End Try

        If remplacerComptes Then EcrireLesComptes(lot, progression, messageErreur)

        Return True
    End Function

    ''' <summary>Les clés que le rapport a marquées « à créer », objet compris.</summary>
    Private Shared Function ClesACreer(rapport As Rapport) As HashSet(Of String)

        Dim cles As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        For Each ligne As LigneRapport In rapport.Lignes
            If ligne.Etat <> EtatLigneWU.ACreer Then Continue For
            cles.Add(ligne.Objet & "|" & ligne.Cle)
        Next

        Return cles
    End Function

    Private Shared Function EcrireLesGroupes(lot As LotParametrageWU, aCreer As HashSet(Of String),
                                             connexion As SqlConnection, transaction As SqlTransaction,
                                             progression As ProgressionWU) As Integer

        Dim ecrits As Integer = 0

        For Each groupe As GroupeStatistiqueWU In lot.Groupes

            If Not aCreer.Contains("Groupe|" & groupe.Nom) Then Continue For

            Annoncer(progression, $"Groupe {groupe.Nom}")
            Porter(DemandeWU.DepuisGroupe(groupe, OperationWU.Creation), "groupe", groupe.Nom,
                   AddressOf PdvRepository.AppliquerGroupe, connexion, transaction)

            ecrits += 1
        Next

        Return ecrits
    End Function

    Private Shared Function EcrireLesAgences(lot As LotParametrageWU, aCreer As HashSet(Of String),
                                             connexion As SqlConnection, transaction As SqlTransaction,
                                             progression As ProgressionWU) As Integer

        Dim ecrits As Integer = 0

        For Each agence As PointDeVenteEC In lot.Agences

            If Not aCreer.Contains("Agence propre|" & agence.CodeSite) Then Continue For

            Annoncer(progression, $"Agence {agence.CodeSite}")
            Porter(DemandeWU.DepuisAgence(agence, OperationWU.Creation), "agence", agence.CodeSite,
                   AddressOf PdvRepository.AppliquerAgence, connexion, transaction)

            ecrits += 1
        Next

        Return ecrits
    End Function

    Private Shared Function EcrireLesSousAgents(lot As LotParametrageWU, aCreer As HashSet(Of String),
                                                connexion As SqlConnection, transaction As SqlTransaction,
                                                progression As ProgressionWU) As Integer

        Dim ecrits As Integer = 0

        For Each pdv As PointDeVenteSA In lot.SousAgents

            If Not aCreer.Contains("Sous-agent|" & pdv.CodePdv) Then Continue For

            Annoncer(progression, $"Sous-agent {pdv.CodePdv}")
            Porter(DemandeWU.DepuisSousAgent(pdv, OperationWU.Creation), "sous-agent", pdv.CodePdv,
                   AddressOf PdvRepository.AppliquerSousAgent, connexion, transaction)

            ecrits += 1
        Next

        Return ecrits
    End Function

    ''' <summary>
    ''' Signature des méthodes d'écriture du référentiel : celles-là mêmes qu'emprunte une
    ''' demande autorisée. L'import n'a donc PAS son propre SQL — il écrit par le chemin que
    ''' le double regard emprunte déjà, et ce qui est vrai de l'un l'est de l'autre.
    ''' </summary>
    Private Delegate Function EcritureReferentiel(demande As DemandeWU, connexion As SqlConnection,
                                                  transaction As SqlTransaction,
                                                  ByRef messageErreur As String) As Boolean

    ''' <summary>
    ''' Porte une ligne, et lève si elle ne passe pas : dans une transaction, une écriture
    ''' refusée doit tout arrêter. Un import qui « continue malgré tout » laisserait un
    ''' référentiel dont personne ne sait ce qu'il contient.
    ''' </summary>
    Private Shared Sub Porter(demande As DemandeWU, objet As String, cle As String,
                              ecrire As EcritureReferentiel,
                              connexion As SqlConnection, transaction As SqlTransaction)

        Dim messageErreur As String = String.Empty

        If ecrire(demande, connexion, transaction, messageErreur) Then Return

        Throw New ApplicationException($"{objet} « {cle} » : {messageErreur}")
    End Sub

    ''' <summary>
    ''' Écrit les comptes comptables, en conservant le CODE DE LA LIGNE LOCALE.
    '''
    ''' Le code désigne la ligne de SystemeWU dans la base d'origine ; la base d'arrivée a la
    ''' sienne, créée par le script d'installation. Reprendre celui du fichier viserait une
    ''' ligne qui n'existe pas, et la mise à jour ne toucherait rien.
    ''' </summary>
    Private Shared Sub EcrireLesComptes(lot As LotParametrageWU, progression As ProgressionWU,
                                        ByRef messageErreur As String)

        If lot.Comptes Is Nothing Then Return

        Annoncer(progression, "Comptes comptables")

        Dim motif As String = String.Empty
        Dim locaux As ComptesSystemeWU = WURepository.ChargerComptesSysteme(motif)

        Dim aEcrire As New ComptesSystemeWU() With {
            .CompteCourant = lot.Comptes.CompteCourant,
            .CompteInterBancaire = lot.Comptes.CompteInterBancaire,
            .CommissionTransfertBanque = lot.Comptes.CommissionTransfertBanque,
            .CommissionEnvoiBanque = lot.Comptes.CommissionEnvoiBanque,
            .CommissionPaiementBanque = lot.Comptes.CommissionPaiementBanque,
            .ImpotsTaxeEnvoi = lot.Comptes.ImpotsTaxeEnvoi,
            .TVACollectee = lot.Comptes.TVACollectee,
            .TTAEnvoi = lot.Comptes.TTAEnvoi,
            .TTAReception = lot.Comptes.TTAReception,
            .CodeParametrage = If(locaux Is Nothing, String.Empty, locaux.CodeParametrage)
        }

        Dim echec As String = String.Empty

        If WURepository.EnregistrerComptesSysteme(aEcrire, echec) Then
            ComptesSystemeWU.Actuels = aEcrire
            Return
        End If

        messageErreur = "Le référentiel est chargé, mais les COMPTES COMPTABLES ne l'ont pas été :" &
                        Environment.NewLine & echec & Environment.NewLine & Environment.NewLine &
                        "Saisissez-les sur l'écran « Comptes systèmes » : ils sont neuf, et la " &
                        "pièce comptable ne s'équilibrera pas sans eux."
    End Sub

    Private Shared Sub Annoncer(progression As ProgressionWU, libelle As String)
        If progression Is Nothing Then Return
        progression.Avancer(libelle)
    End Sub

#End Region

End Class
