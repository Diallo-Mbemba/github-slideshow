Option Strict On
Option Explicit On

Imports System.Globalization

''' <summary>
''' L'en-tête du TRAITEMENT d'une journée, là où la pièce comptable porte les écritures.
'''
''' POURQUOI UN DOCUMENT DE PLUS
'''
''' La pièce porte déjà les quatre cartouches de la banque, et les écritures sont donc
''' couvertes par une signature. Mais qui signe la pièce signe ce qui y figure — et cinq
''' choses n'y figurent pas : ce qui n'a PAS été comptabilisé, quels rapports Western Union
''' ont servi, l'écart d'arrondi isolé, le fichier destiné au core banking, et la preuve que
''' la journée est complète.
'''
''' Cet en-tête porte ces cinq choses. Le bordereau qui s'en tire atteste de la FAÇON dont la
''' journée a été faite, là où la pièce atteste de ce qui a été écrit.
'''
''' LE VISA
'''
''' Modifier un taux de sous-agent exige deux personnes ; comptabiliser une journée entière
''' n'en exigeait qu'une. Le visa corrige ce déséquilibre : un supérieur relit la journée et
''' la marque visée, et ce n'est pas lui qui l'a comptabilisée — la base le refuse.
'''
''' Recomptabiliser une journée EFFACE son visa. Le visa atteste d'un traitement précis ;
''' refaire la journée en produit un autre, et laisser le visa en place ferait croire qu'un
''' supérieur a vu des chiffres qu'il n'a jamais vus.
''' </summary>
Public Class TraitementJourneeWU

#Region "La journée"

    ''' <summary>
    ''' La journée sous laquelle la pièce est classée : LE PREMIER JOUR COUVERT par le rapport,
    ''' c'est-à-dire le plus ancien. Elle est la clé de l'historique, de la pièce et de cet
    ''' en-tête.
    ''' </summary>
    Public Property DateActivite As Date

    Public Property DateValeur As Date?
    Public Property NumeroLot As String = String.Empty

    ''' <summary>
    ''' LA PÉRIODE RÉELLEMENT COUVERTE par les rapports chargés — premier et dernier jour.
    '''
    ''' POURQUOI ELLE EST CONSERVÉE, ALORS QUE DateActivite SUFFIT À CLASSER LA PIÈCE. La
    ''' banque liquide à la semaine : un rapport porte six journées, et l'en-tête n'en gardait
    ''' qu'une. L'avertissement affiché avant de comptabiliser — « ne chargez pas ensuite un
    ''' rapport d'une journée déjà comprise dans cette période » — était donc une promesse que
    ''' le code ne pouvait pas tenir : recharger le 24 seul après une semaine du 24 au 30 ne
    ''' trouvait rien et passait sans un mot.
    '''
    ''' ELLES NE SONT PAS CLÉ PRIMAIRE, et ne le deviendront pas. DateActivite l'est pour
    ''' T_HistoriqueWU, T_PieceWU et T_TraitementWU, plus leurs archives : changer cela
    ''' demanderait une reprise de toute la base en production pour un gain nul — une période
    ''' est désignée sans ambiguïté par son premier jour, dès lors que deux périodes qui se
    ''' chevauchent ne peuvent plus être comptabilisées l'une après l'autre sans avertissement.
    '''
    ''' VIDES POUR UNE JOURNÉE DÉJÀ COMPTABILISÉE avant cette version : rien ne les reconstitue
    ''' rétroactivement, et le contrôle de chevauchement retombe alors sur la seule DateActivite.
    ''' </summary>
    Public Property DebutPeriode As Date?
    Public Property FinPeriode As Date?

    ''' <summary>
    ''' Le premier et le dernier jour couverts, pour qui veut les lire sans se demander si la
    ''' période a été enregistrée. Sur un en-tête d'avant cette version, les deux valent la
    ''' date d'activité : une journée seule est une période d'un jour.
    ''' </summary>
    Public ReadOnly Property PremierJour As Date
        Get
            If DebutPeriode.HasValue Then Return DebutPeriode.Value.Date
            Return DateActivite.Date
        End Get
    End Property

    Public ReadOnly Property DernierJour As Date
        Get
            If FinPeriode.HasValue Then Return FinPeriode.Value.Date
            Return PremierJour
        End Get
    End Property

#End Region

#Region "Les sources"

    ''' <summary>Nom du rapport d'activité tel qu'il a été chargé.</summary>
    Public Property FichierActivite As String = String.Empty

    ''' <summary>
    ''' Empreinte SHA-256 du rapport d'activité.
    '''
    ''' Elle ne protège de rien — qui remplace un fichier peut recalculer la sienne. Elle
    ''' répond à une question : « le fichier que vous me montrez est-il celui qui a été traité
    ''' ce jour-là ? ». Sans elle, la question n'a pas de réponse.
    ''' </summary>
    Public Property EmpreinteActivite As String = String.Empty

    Public Property FichierReglement As String = String.Empty
    Public Property EmpreinteReglement As String = String.Empty

#End Region

#Region "Ce qui a été traité, et ce qui ne l'a pas été"

    Public Property NombrePdv As Integer = 0
    Public Property NombreSousAgents As Integer = 0
    Public Property NombreAgences As Integer = 0

    ''' <summary>
    ''' Accounts non paramétrés, donc écartés de la comptabilisation. Ils n'apparaissent
    ''' nulle part sur la pièce — c'est précisément pour eux que le bordereau existe.
    ''' </summary>
    Public Property NombreEcartes As Integer = 0

    Public Property NombreEnvois As Integer = 0
    Public Property NombrePaiements As Integer = 0
    Public Property NombreAnnulations As Integer = 0

#End Region

#Region "La pièce, en trois chiffres"

    Public Property TotalDebit As Long = 0L
    Public Property TotalCredit As Long = 0L

    ''' <summary>Écart d'arrondi posé sur le compte inter bancaire, signé.</summary>
    Public Property EcartArrondi As Long = 0L
    Public Property CompteEcart As String = String.Empty

    Public ReadOnly Property Equilibree As Boolean
        Get
            Return TotalDebit = TotalCredit
        End Get
    End Property

#End Region

#Region "Qui a fait quoi"

    Public Property ComptabilisePar As String = String.Empty
    Public Property DateComptabilisation As Date?

    Public Property VisePar As String = String.Empty
    Public Property DateVisa As Date?
    Public Property CommentaireVisa As String = String.Empty

    Public ReadOnly Property EstVisee As Boolean
        Get
            Return VisePar.Length > 0 AndAlso DateVisa.HasValue
        End Get
    End Property

    ''' <summary>
    ''' Vrai si cet en-tête a été retrouvé en base. Faux pour une journée comptabilisée avant
    ''' la mise en service de la table : son bordereau se reconstitue alors depuis l'historique
    ''' et la pièce, mais sans le nom des rapports, qui n'était conservé nulle part.
    ''' </summary>
    Public Property Enregistre As Boolean = False

#End Region

#Region "Libellés"

    ''' <summary>L'état du visa, en une phrase.</summary>
    Public ReadOnly Property LibelleVisa As String
        Get
            If Not EstVisee Then Return "EN ATTENTE DE VISA"

            Return $"Visée le {DateVisa.Value.ToString("dd/MM/yyyy à HH:mm", CultureInfo.InvariantCulture)}" &
                   $" par {VisePar}"
        End Get
    End Property

    ''' <summary>Le traitement en une phrase, pour un journal ou une confirmation.</summary>
    Public ReadOnly Property Intitule As String
        Get
            Return $"Journée du {DateActivite.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture)}" &
                   If(NumeroLot.Length > 0, $" — lot {NumeroLot}", String.Empty)
        End Get
    End Property

#End Region

End Class
