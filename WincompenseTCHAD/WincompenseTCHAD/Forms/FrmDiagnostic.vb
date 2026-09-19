Option Strict On
Option Explicit On

Imports System.Windows.Forms

''' <summary>
''' Présente un diagnostic long — celui de <see cref="DiagnosticSqlWU"/> en pratique — que
''' les libellés des écrans ne peuvent pas contenir.
'''
''' Le bouton « Copier » n'est pas un ornement : le texte contient le T-SQL à exécuter sur le
''' serveur, et c'est l'agent qui le transmettra à l'informatique. Le lui faire recopier à la
''' main reviendrait à lui faire inventer un nom de compte.
''' </summary>
Public Class FrmDiagnostic

    ''' <summary>
    ''' Affiche un diagnostic au-dessus de la fenêtre appelante. Ne fait rien si le texte est
    ''' vide : un écran de diagnostic sans diagnostic inquiéterait pour rien.
    ''' </summary>
    Public Shared Sub Afficher(proprietaire As IWin32Window, titre As String, texte As String)

        If String.IsNullOrWhiteSpace(texte) Then Return

        Using fenetre As New FrmDiagnostic()

            fenetre.lblTitre.Text = If(String.IsNullOrWhiteSpace(titre), "Diagnostic", titre)
            fenetre.Text = fenetre.lblTitre.Text

            ' Les sauts de ligne simples d'un message composé en vbLf ne s'afficheraient pas
            ' dans une zone de texte Windows, qui attend vbCrLf.
            fenetre.txtDiagnostic.Text = Normaliser(texte)
            fenetre.txtDiagnostic.Select(0, 0)

            If proprietaire Is Nothing Then
                fenetre.StartPosition = FormStartPosition.CenterScreen
                fenetre.ShowDialog()
            Else
                fenetre.ShowDialog(proprietaire)
            End If
        End Using
    End Sub

    ''' <summary>Ramène tous les sauts de ligne à la forme attendue par Windows.</summary>
    Private Shared Function Normaliser(texte As String) As String
        Return texte.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf).Replace(vbLf, vbCrLf)
    End Function

    Private Sub btnCopier_Click(sender As Object, e As EventArgs) Handles btnCopier.Click

        Try
            Clipboard.SetText(txtDiagnostic.Text)
            btnCopier.Text = "Copié"

        Catch ex As System.Runtime.InteropServices.ExternalException
            ' Presse-papiers verrouillé par une autre application : le texte reste lisible
            ' et sélectionnable à l'écran, rien n'est perdu.
            MessageBox.Show(Me,
                            "Le presse-papiers est occupé par une autre application." & vbCrLf & vbCrLf &
                            "Sélectionnez le texte à la souris, puis Ctrl+C.",
                            "Copie impossible", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Try
    End Sub

    Private Sub btnFermer_Click(sender As Object, e As EventArgs) Handles btnFermer.Click
        Close()
    End Sub

End Class
