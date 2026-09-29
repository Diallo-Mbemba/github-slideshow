Option Strict On
Option Explicit On

Imports System.Windows.Forms

''' <summary>Point d'entrée de l'application Windows Forms.</summary>
Module Program

    ''' <summary>
    ''' L'enchaînement du démarrage : s'identifier une fois, puis choisir un produit de
    ''' transfert, le traiter, et revenir au choix autant de fois qu'on le veut.
    '''
    ''' POURQUOI UNE BOUCLE, ET NON UN SEUL Application.Run
    '''
    ''' La version précédente faisait « Application.Run(New FrmPrincipal()) » : fermer
    ''' l'espace de travail arrêtait l'application. Changer de produit aurait donc voulu
    ''' dire quitter et se reconnecter. La boucle sépare les deux : l'espace de travail se
    ''' ferme soit pour revenir au choix, soit pour quitter, et c'est LUI qui le dit.
    '''
    ''' L'identification, elle, reste HORS de la boucle : changer de produit ne change pas
    ''' d'utilisateur. Un utilisateur habilité à traiter la compense la traite pour tous les
    ''' produits — le produit n'est pas un droit.
    ''' </summary>
    <STAThread()>
    Sub Main()
        Application.EnableVisualStyles()
        Application.SetCompatibleTextRenderingDefault(False)

        ' L'identification précède tout : ni le choix du produit ni la fenêtre principale ne
        ' sont construits tant que personne n'est connecté. Les écrans interrogent SessionWU
        ' pour connaître leurs droits, et cette classe ne répond rien tant qu'aucune session
        ' n'est ouverte.
        If Not Identifier() Then Return

        ' Les produits déclarés par la banque sont lus UNE FOIS, après l'identification et
        ' avant le premier choix : la fenêtre de choix doit montrer ce que la banque a
        ' déclaré, et non la liste interne. Si la table n'existe pas encore, le dépôt installe
        ' la liste de secours et l'application démarre quand même.
        Dim messageErreur As String = String.Empty
        ProduitRepository.Charger(messageErreur)

        Try
            Do
                Dim produit As ProduitTransfert = ChoisirLeProduit()
                If produit Is Nothing Then Return

                ProduitTransfert.Actif = produit

                ' Le produit est posé AVANT la construction de l'espace de travail : son titre
                ' le nomme et ses écrans en dépendent dès l'ouverture.
                Dim espace As New FrmPrincipal()
                Application.Run(espace)

                ' La propriété se lit après la fermeture — Application.Run a libéré la
                ' fenêtre, mais ce drapeau est un simple booléen : il ne touche à aucune
                ' ressource libérée.
                If Not espace.RetourAuChoixDuProduit Then Return

                ' Relecture à chaque retour : la banque vient peut-être d'en déclarer un
                ' depuis l'espace qu'elle quitte, et il doit apparaître sans relancer
                ' l'application.
                ProduitRepository.Charger(messageErreur)
            Loop

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

    ''' <summary>
    ''' Affiche la fenêtre de choix du produit de transfert. Retourne Nothing si
    ''' l'utilisateur quitte : l'application s'arrête alors, la session déjà ouverte étant
    ''' refermée par le Finally de Main.
    ''' </summary>
    Private Function ChoisirLeProduit() As ProduitTransfert

        Using choix As New FrmChoixProduit()

            If choix.ShowDialog() <> DialogResult.OK Then Return Nothing
            Return choix.ProduitChoisi
        End Using
    End Function

End Module
