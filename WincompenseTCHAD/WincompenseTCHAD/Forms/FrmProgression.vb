Option Strict On
Option Explicit On

Imports System.Windows.Forms

''' <summary>
''' Fenêtre d'avancement d'une opération longue : ce qui se fait, et où l'on en est.
'''
''' COMMENT ELLE SE RAFRAÎCHIT, ET POURQUOI AINSI
'''
''' Excel se pilote depuis le fil de l'interface, et lui seul. L'opération occupe donc ce fil
''' du début à la fin : la fenêtre ne peut pas se redessiner d'elle-même, et un fil
''' d'arrière-plan ferait échouer les appels COM.
'''
''' Elle se redessine donc SUR ORDRE, à chaque étape annoncée, par Refresh(). Ce n'est pas
''' DoEvents : Refresh repeint, il ne dépile pas les messages en attente. Un second clic sur le
''' bouton d'export ne peut donc pas se glisser au milieu d'un export en cours — c'est
''' exactement ce qu'on veut, et c'est aussi pourquoi il n'y a pas de bouton Annuler.
'''
''' Entre deux étapes — pendant l'écriture d'un onglet — Windows peut griser la fenêtre. C'est
''' le prix du pilotage d'Excel, et la raison pour laquelle les étapes sont annoncées souvent.
'''
''' ELLE N'A NI BOUTON NI CROIX. Elle se ferme d'elle-même, et couvre le bouton qui l'a
''' ouverte : le vrai risque d'une opération longue est le double-clic qui lance deux Excel.
''' </summary>
Public Class FrmProgression

    ''' <summary>
    ''' Ce que les services reçoivent. La fenêtre écoute, les services ne la connaissent pas.
    ''' </summary>
    Public ReadOnly Property Progression As ProgressionWU
        Get
            Return _progression
        End Get
    End Property
    Private ReadOnly _progression As New ProgressionWU()

    Public Sub New()

        InitializeComponent()
        IconesWU.Habiller(Me)

        AddHandler _progression.EtapeChangee, AddressOf SurEtape
        AddHandler _progression.AvancementChange, AddressOf SurAvancement
    End Sub

#Region "Ouverture et fermeture"

    ''' <summary>
    ''' Ouvre la fenêtre au-dessus de l'écran appelant et rend la main tout de suite : c'est
    ''' l'appelant qui travaille ensuite, et qui la ferme.
    '''
    ''' Elle s'affiche SANS DÉLAI. Un affichage différé — pour éviter un clignotement sur les
    ''' opérations brèves — la ferait justement manquer dans le seul cas qui compte : une
    ''' étape unique et longue, où aucun événement n'arrive pendant qu'on attend.
    ''' </summary>
    ''' <param name="proprietaire">Écran appelant, au-dessus duquel se centrer.</param>
    ''' <param name="titre">Ce qui se passe, en une ligne.</param>
    Public Shared Function Ouvrir(proprietaire As IWin32Window, titre As String) As FrmProgression

        Dim fenetre As New FrmProgression()

        fenetre.lblTitre.Text = "   " & If(titre, "Traitement en cours")
        fenetre.Text = If(titre, "Traitement en cours")

        Dim parente As Form = TryCast(proprietaire, Form)

        If parente Is Nothing Then
            fenetre.StartPosition = FormStartPosition.CenterScreen
            fenetre.Show()
        Else
            fenetre.Show(parente)
        End If

        fenetre.Refresh()
        Return fenetre
    End Function

    ''' <summary>
    ''' Ferme la fenêtre. Sans effet si elle l'est déjà : l'appelant la ferme dans son Finally,
    ''' et une erreur ne doit pas en produire une seconde.
    ''' </summary>
    Public Sub Fermer()

        Try
            If Not IsDisposed Then Close()
        Catch ex As ObjectDisposedException
            ' Déjà fermée : rien à faire.
        End Try
    End Sub

#End Region

#Region "Réception des étapes"

    Private Sub SurEtape(libelle As String)

        If IsDisposed Then Return

        lblEtape.Text = libelle
        Rafraichir()
    End Sub

    Private Sub SurAvancement(rang As Integer, total As Integer)

        If IsDisposed Then Return

        If total <= 0 Then

            ' Total inconnu : une barre immobile à zéro ferait croire à un blocage, et une barre
            ' en défilement ne défilerait pas — son animation demande un fil libre, que
            ' l'opération occupe. On la retire, et le libellé porte seul l'information.
            barre.Visible = False
            lblCompteur.Text = String.Empty
        Else
            barre.Visible = True
            barre.Maximum = total
            barre.Value = Math.Min(Math.Max(rang, 0), total)
            lblCompteur.Text = $"Étape {barre.Value} sur {total}"
        End If

        Rafraichir()
    End Sub

    ''' <summary>
    ''' Repeint immédiatement. Refresh et non DoEvents : voir la remarque en tête de classe.
    ''' </summary>
    Private Sub Rafraichir()

        Try
            Refresh()
        Catch ex As ObjectDisposedException
            ' La fenêtre s'est fermée entre-temps : l'opération, elle, continue.
        End Try
    End Sub

#End Region

End Class
