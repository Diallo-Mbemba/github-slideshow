Option Strict On
Option Explicit On

Imports System.Globalization

''' <summary>
''' Le paramétrage d'une base, sorti dans un fichier ou relu depuis un fichier.
'''
''' CE QU'IL CONTIENT, ET POURQUOI CELA S'ARRÊTE LÀ
'''
''' Les comptes comptables, les groupes statistiques, les sous-agents et les agences propres :
''' tout ce qu'il faudrait resaisir à la main pour remettre une installation en marche, et
''' rien d'autre.
'''
''' Il ne contient NI l'historique, NI les pièces, NI les demandes. Ce n'est pas une
''' sauvegarde de la base — la sauvegarde d'une base SQL Server se fait sur le serveur, par
''' l'équipe qui le tient, et aucune application cliente ne peut la remplacer. C'est un
''' filet : de quoi ne pas repartir d'une feuille blanche.
'''
''' LES UTILISATEURS SONT LÀ POUR MÉMOIRE, ET NE SE RECHARGENT PAS
'''
''' Leur liste est exportée — identifiant, nom, rôle, fonction — pour qu'on sache QUI
''' recréer. Leurs mots de passe, eux, ne quittent jamais la base, pas même sous forme
''' d'empreinte : un fichier qui circule ne porte pas les identifiants d'une banque. Les
''' recharger supposerait soit d'emporter ces empreintes, soit d'inventer des mots de passe
''' provisoires et de les écrire dans le même fichier. Ni l'un ni l'autre.
''' </summary>
Public Class LotParametrageWU

#Region "D'où il vient"

    ''' <summary>Version du format de fichier. Un fichier plus récent que l'application est refusé.</summary>
    Public Property Version As String = String.Empty

    Public Property Base As String = String.Empty
    Public Property Serveur As String = String.Empty
    Public Property DateExport As Date
    Public Property ExportePar As String = String.Empty

    ''' <summary>
    ''' Vrai si le contenu ne correspond plus à l'empreinte inscrite dans le manifeste.
    '''
    ''' Ce n'est pas une erreur : corriger un taux dans le fichier avant de le recharger est
    ''' légitime, et parfois la seule façon de s'en sortir. Mais celui qui charge doit le
    ''' SAVOIR, parce qu'il ne charge alors plus ce que l'application avait produit.
    ''' </summary>
    Public Property Modifie As Boolean = False

#End Region

#Region "Ce qu'il porte"

    Public Property Comptes As ComptesSystemeWU = Nothing

    Public Property Groupes As New List(Of GroupeStatistiqueWU)()
    Public Property SousAgents As New List(Of PointDeVenteSA)()
    Public Property Agences As New List(Of PointDeVenteEC)()

    ''' <summary>Pour information seulement : cette liste n'est jamais réécrite dans une base.</summary>
    Public Property Utilisateurs As New List(Of UtilisateurWU)()

#End Region

#Region "Libellés"

    ''' <summary>Ce que le lot contient, en une phrase.</summary>
    Public ReadOnly Property Intitule As String
        Get
            Return $"{Groupes.Count} groupe(s), {SousAgents.Count} sous-agent(s), " &
                   $"{Agences.Count} agence(s) propre(s)" &
                   If(Comptes Is Nothing, String.Empty, ", les comptes comptables")
        End Get
    End Property

    ''' <summary>D'où vient ce lot, en une phrase.</summary>
    Public ReadOnly Property Provenance As String
        Get
            Dim phrase As String = $"Exporté le {DateExport.ToString("dd/MM/yyyy à HH:mm", CultureInfo.InvariantCulture)}"

            If ExportePar.Length > 0 Then phrase &= $" par {ExportePar}"
            If Base.Length > 0 Then phrase &= $", base {Base}"
            If Serveur.Length > 0 Then phrase &= $" sur {Serveur}"

            Return phrase & "."
        End Get
    End Property

    ''' <summary>Vrai si le lot ne porte rien : un fichier vide n'a rien à charger.</summary>
    Public ReadOnly Property EstVide As Boolean
        Get
            Return Comptes Is Nothing AndAlso Groupes.Count = 0 AndAlso
                   SousAgents.Count = 0 AndAlso Agences.Count = 0
        End Get
    End Property

#End Region

End Class
