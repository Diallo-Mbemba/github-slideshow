Option Strict On
Option Explicit On

''' <summary>
''' Une fiche du réseau Western Union, telle qu'elle s'imprime : un sous-agent ou une agence
''' propre, avec ce qu'on sait de sa création.
'''
''' D'OÙ VIENNENT LA DATE ET L'AUTEUR. Ni T_Pdv_SA ni T_Pdv_EC ne portent de date de création
''' ni de créateur : ces deux tables sont le référentiel de la banque, et n'ont jamais eu ces
''' colonnes. Les deux valeurs sont donc RECONSTITUÉES depuis T_DemandeWU, où le circuit à
''' quatre yeux garde trace de chaque création autorisée.
'''
''' ELLES PEUVENT MANQUER, ET C'EST NORMAL. Une fiche antérieure au circuit, ou créée
''' directement par un gestionnaire, n'a aucune demande : sa date et son auteur sont alors
''' INCONNUS, et c'est ce que l'état écrit. Inventer une valeur — « référentiel banque »,
''' la date du jour, le premier administrateur venu — reviendrait à affirmer dans un document
''' imprimé quelque chose que la base ne dit pas.
''' </summary>
Public Class LigneReseauWU

    ''' <summary>Ce qu'affiche la colonne « Créé le » quand aucune demande n'a été trouvée.</summary>
    Public Const DATE_INCONNUE As String = "inconnue"

    ''' <summary>Ce qu'affiche la colonne « Créé par » quand aucune demande n'a été trouvée.</summary>
    Public Const AUTEUR_INCONNU As String = "(inconnu)"

    ''' <summary>Account : Code_Pdv pour un sous-agent, Codesite pour une agence propre.</summary>
    Public Property Account As String = String.Empty

    ''' <summary>Désignation de l'agence.</summary>
    Public Property Designation As String = String.Empty

    ''' <summary>Groupe statistique. Toujours vide pour une agence propre : elle n'en a pas.</summary>
    Public Property Groupe As String = String.Empty

    ''' <summary>Code agence : codeagence pour un sous-agent, [CodeAgenc-Voyager] pour une agence.</summary>
    Public Property CodeAgence As String = String.Empty

    ''' <summary>Vrai pour un sous-agent, faux pour une agence propre.</summary>
    Public Property EstSousAgent As Boolean = True

    ''' <summary>Date de création, ou Nothing si aucune demande autorisée ne la porte.</summary>
    Public Property DateCreation As Date?

    ''' <summary>Identifiant de celui qui a déposé la demande de création. Vide si inconnu.</summary>
    Public Property CreePar As String = String.Empty

    ''' <summary>Vrai quand la création de cette fiche a pu être datée.</summary>
    Public ReadOnly Property DateConnue As Boolean
        Get
            Return DateCreation.HasValue
        End Get
    End Property

    ''' <summary>La date telle qu'elle s'affiche et s'imprime.</summary>
    Public ReadOnly Property DateLisible As String
        Get
            If Not DateCreation.HasValue Then Return DATE_INCONNUE
            Return DateCreation.Value.ToString("dd/MM/yyyy", Globalization.CultureInfo.InvariantCulture)
        End Get
    End Property

    ''' <summary>L'auteur tel qu'il s'affiche et s'imprime.</summary>
    Public ReadOnly Property AuteurLisible As String
        Get
            Dim auteur As String = If(CreePar, String.Empty).Trim()
            Return If(auteur.Length = 0, AUTEUR_INCONNU, auteur)
        End Get
    End Property

End Class

''' <summary>
''' Une section de l'état : son titre, et les fiches qu'elle totalise.
'''
''' Le titre porte tout ce qui caractérise la section — le groupe, son effectif, ses comptes,
''' son taux — parce que c'est lui qui coiffe le tableau sur le papier. Il est composé une
''' fois par ListeReseauService, et non reconstruit à chaque affichage : l'écran et le document
''' imprimé doivent nommer la section exactement de la même façon.
''' </summary>
Public Class SectionReseauWU

    Public Property Titre As String = String.Empty

    Public Property Lignes As New List(Of LigneReseauWU)

    ''' <summary>Vrai pour la section des agences propres : elle ne compte pas de sous-agents.</summary>
    Public Property EstSectionDAgences As Boolean = False

    Public ReadOnly Property Nombre As Integer
        Get
            Return Lignes.Count
        End Get
    End Property

End Class
