Option Strict On
Option Explicit On

Imports System.Windows.Forms

''' <summary>Point d'entrée de l'application Windows Forms.</summary>
Module Program

    <STAThread()>
    Sub Main()
        Application.EnableVisualStyles()
        Application.SetCompatibleTextRenderingDefault(False)

        ' L'identification précède tout : la fenêtre principale n'est même pas construite tant
        ' que personne n'est connecté. Les écrans interrogent SessionWU pour connaître leurs
        ' droits, et cette classe ne répond rien tant qu'aucune session n'est ouverte.
        If Not Identifier() Then Return

        Try
            Application.Run(New FrmPrincipal())
        Finally
            SessionWU.Fermer()
        End Try
    End Sub

    ''' <summary>
    ''' Affiche l'écran de connexion et ouvre la session. Retourne faux si l'utilisateur
    ''' abandonne : l'application s'arrête alors sans rien afficher d'autre.
    ''' </summary>
    Private Function Identifier() As Boolean

        Using connexion As New FrmConnexion()

            If connexion.ShowDialog() <> DialogResult.OK OrElse connexion.UtilisateurConnecte Is Nothing Then
                Return False
            End If

            SessionWU.Ouvrir(connexion.UtilisateurConnecte)
            Return True
        End Using
    End Function

End Module
