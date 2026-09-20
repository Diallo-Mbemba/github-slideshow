Option Strict On
Option Explicit On

Imports System.Windows.Forms

''' <summary>
''' Réglages de grilles partagés par les écrans.
'''
''' Une seule règle y vit pour l'instant, mais elle vaut d'être écrite une fois : poser une
''' largeur de colonne sans l'avoir d'abord sortie du redimensionnement automatique fait
''' tomber le moteur de disposition de Windows Forms, et l'erreur remonte loin de sa cause.
''' </summary>
Public NotInheritable Class GrilleWU

    Private Sub New()
    End Sub

    ''' <summary>
    ''' Fixe la largeur d'une colonne, en la sortant D'ABORD du redimensionnement automatique.
    '''
    ''' POURQUOI CE DÉTOUR. Une grille en AutoSizeColumnsMode.Fill partage la largeur
    ''' disponible entre ses colonnes : leur Width ne leur appartient plus, et y écrire une
    ''' valeur fait tomber le calcul de disposition. Le symptôme est une exception sur la ligne
    ''' d'affectation, qui ne dit rien de la grille ni du mode en cause.
    '''
    ''' Le remède tient en une ligne — AutoSizeMode = None sur la colonne — mais c'est une
    ''' ligne qu'on oublie. En faire un sous-programme, et refuser tout Width posé autrement,
    ''' vaut mieux que de compter sur la mémoire.
    '''
    ''' La colonne est ignorée si elle n'existe pas : une grille dont le jeu de colonnes
    ''' dépend des données ne doit pas imposer un test à chaque appelant.
    ''' </summary>
    Public Shared Sub LargeurFixe(grille As DataGridView, nomColonne As String, largeur As Integer)

        If grille Is Nothing OrElse String.IsNullOrEmpty(nomColonne) Then Return
        If Not grille.Columns.Contains(nomColonne) Then Return

        Dim colonne As DataGridViewColumn = grille.Columns(nomColonne)
        If colonne Is Nothing Then Return

        colonne.AutoSizeMode = DataGridViewAutoSizeColumnMode.None

        ' La largeur minimale prime sur la largeur demandée : une colonne étroite resterait
        ' large si son minimum n'était pas abaissé d'abord.
        colonne.MinimumWidth = Math.Min(largeur, colonne.MinimumWidth)
        colonne.Width = largeur
    End Sub

End Class
