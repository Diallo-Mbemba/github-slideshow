Option Strict On
Option Explicit On

''' <summary>
''' Rend compte de l'avancement d'une opération longue, sans rien savoir de l'écran.
'''
''' POURQUOI UN CONTRAT PLUTÔT QU'UNE FENÊTRE
'''
''' Les services qui écrivent dans Excel n'ont pas à connaître Windows Forms. S'ils recevaient
''' une fenêtre, ils deviendraient inappelables ailleurs — depuis un traitement de nuit, ou
''' depuis un test. Ils reçoivent donc un objet qui ÉMET, et c'est la fenêtre qui écoute.
'''
''' LE PARAMÈTRE EST TOUJOURS FACULTATIF côté service. Sans lui, le service se comporte
''' exactement comme avant : un appel qui ne veut rien afficher n'a rien à passer, et aucun
''' appel existant ne casse.
'''
''' CE QUE CET OBJET NE FAIT PAS
'''
''' Il n'annule rien. Excel se pilote depuis le fil de l'interface et lui seul : pour recevoir
''' un clic pendant l'écriture, il faudrait pomper les messages Windows au milieu d'une
''' séquence d'appels COM — c'est-à-dire autoriser un second export à démarrer pendant le
''' premier. Et une interruption laisserait Excel ouvert sur un classeur à moitié écrit, que
''' personne ne saurait distinguer d'un classeur complet.
'''
''' Il informe. Il ne négocie pas.
''' </summary>
Public NotInheritable Class ProgressionWU

#Region "Événements"

    ''' <summary>Une nouvelle étape commence. Le libellé est destiné à être lu tel quel.</summary>
    Public Event EtapeChangee(libelle As String)

    ''' <summary>L'avancement a bougé : rang atteint sur un total. Total à zéro : indéterminé.</summary>
    Public Event AvancementChange(rang As Integer, total As Integer)

#End Region

#Region "État"

    ''' <summary>Nombre d'étapes annoncé. Zéro tant qu'on ne le connaît pas.</summary>
    Public ReadOnly Property Total As Integer
        Get
            Return _total
        End Get
    End Property
    Private _total As Integer = 0

    ''' <summary>Étapes franchies.</summary>
    Public ReadOnly Property Rang As Integer
        Get
            Return _rang
        End Get
    End Property
    Private _rang As Integer = 0

    ''' <summary>Dernier libellé annoncé.</summary>
    Public ReadOnly Property Libelle As String
        Get
            Return _libelle
        End Get
    End Property
    Private _libelle As String = String.Empty

#End Region

#Region "Ce que les services appellent"

    ''' <summary>
    ''' Annonce le nombre d'étapes à venir et remet le compteur à zéro.
    '''
    ''' Un total inconnu — zéro — n'est pas une erreur : la barre se met alors en mouvement
    ''' continu, ce qui dit « cela travaille » sans prétendre dire « il reste tant ».
    ''' </summary>
    Public Sub Commencer(total As Integer)

        _total = If(total > 0, total, 0)
        _rang = 0

        RaiseEvent AvancementChange(_rang, _total)
    End Sub

    ''' <summary>Annonce ce qui se fait maintenant, sans avancer le compteur.</summary>
    Public Sub Etape(libelle As String)

        _libelle = If(libelle, String.Empty)
        RaiseEvent EtapeChangee(_libelle)
    End Sub

    ''' <summary>
    ''' Annonce une étape ET la franchit. La forme courante : le service dit ce qu'il fait au
    ''' moment où il le fait, et la barre suit.
    ''' </summary>
    Public Sub Avancer(libelle As String)
        Etape(libelle)
        Avancer()
    End Sub

    ''' <summary>Franchit une étape. Le compteur ne dépasse jamais le total annoncé.</summary>
    Public Sub Avancer()

        If _total > 0 AndAlso _rang >= _total Then
            ' Plus d'étapes franchies qu'annoncées : le total était faux. On ne recule pas la
            ' barre et on ne lève rien — une barre trop courte n'a jamais perdu de données.
            Return
        End If

        _rang += 1
        RaiseEvent AvancementChange(_rang, _total)
    End Sub

    ''' <summary>Porte le compteur à son total : l'opération est terminée.</summary>
    Public Sub Terminer()

        If _total > 0 Then _rang = _total
        RaiseEvent AvancementChange(_rang, _total)
    End Sub

#End Region

End Class
