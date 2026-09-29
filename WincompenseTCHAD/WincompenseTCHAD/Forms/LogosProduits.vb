Option Strict On
Option Explicit On

Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Globalization
Imports System.IO

''' <summary>
''' Le logo d'un produit de transfert : le fichier officiel s'il est là, un emblème dessiné
''' sinon.
'''
''' POURQUOI LE LOGO N'EST PAS DANS L'EXÉCUTABLE
'''
''' « Western Union » et « Ria » sont des marques déposées. Leurs logos ne s'attrapent pas sur
''' une banque d'images — les aperçus y sont filigranés et sous licence, et une licence de
''' photo ne donne aucun droit sur une marque. La banque, elle, est agent agréé : elle a reçu
''' ces logos de Western Union, et les recevra de Ria. Ce sont CES fichiers-là qui doivent
''' s'afficher, et personne d'autre que la banque ne peut les fournir.
'''
''' Ils sont donc LUS À CÔTÉ DE L'EXÉCUTABLE et non incorporés : le jour où le service
''' marketing envoie le logo, on dépose un fichier, sans rien recompiler ni redéployer. C'est
''' aussi ce qui permet à chaque filiale d'avoir les siens.
'''
'''     ...\Wincompense\Logos\WU.png
'''     ...\Wincompense\Logos\RIA.png
'''
''' PNG, JPG, BMP ou GIF ; le PNG à fond transparent rend le mieux. L'image est redimensionnée
''' en conservant ses proportions, et centrée : un logo large ou carré s'affiche correctement
''' sans être déformé.
'''
''' EN ATTENDANT, UN EMBLÈME EST DESSINÉ
'''
''' Un carré arrondi aux couleurs du produit, portant ses initiales. Ce n'est PAS une imitation
''' du logo de la marque, et cela ne doit pas le devenir : c'est une pastille de remplacement,
''' qui se voit comme telle et disparaît dès que le vrai fichier est là.
''' </summary>
Public NotInheritable Class LogosProduits

    Private Sub New()
    End Sub

    ''' <summary>Dossier des logos, cherché à côté de l'exécutable.</summary>
    Public Const DOSSIER As String = "Logos"

    ''' <summary>
    ''' Extensions acceptées, dans l'ordre d'essai. Le PNG d'abord : c'est le seul de la liste
    ''' qui porte la transparence, et un logo sur fond blanc dans un cadre gris se remarque.
    ''' </summary>
    Private Shared ReadOnly EXTENSIONS As String() = {".png", ".jpg", ".jpeg", ".bmp", ".gif"}

    ''' <summary>
    ''' Les logos déjà produits, rangés par produit et par taille.
    '''
    ''' Comme pour les icônes, le cache n'est jamais vidé et ses images ne sont jamais
    ''' libérées : elles sont posées sur des listes et des cadres qui les gardent, et libérer
    ''' une image encore affichée lèverait une exception au premier repaint.
    '''
    ''' Conséquence à connaître : un fichier de logo déposé pendant que l'application tourne
    ''' n'apparaît qu'au lancement suivant.
    ''' </summary>
    Private Shared ReadOnly _cache As New Dictionary(Of String, Bitmap)()

#Region "Accès"

    ''' <summary>
    ''' Rend le logo d'un produit à la taille demandée : le fichier officiel s'il existe et
    ''' qu'il se lit, l'emblème dessiné sinon.
    ''' </summary>
    ''' <param name="produit">Le produit ; Nothing rend Nothing.</param>
    ''' <param name="taille">Côté de l'image en pixels.</param>
    Public Shared Function Obtenir(produit As ProduitTransfert, taille As Integer) As Bitmap

        If produit Is Nothing Then Return Nothing

        ' Une taille aberrante ne doit pas faire tomber l'écran qui la demande.
        Dim cote As Integer = taille
        If cote < 16 Then cote = 16
        If cote > 512 Then cote = 512

        Dim cle As String = produit.Code & "|" & cote.ToString(CultureInfo.InvariantCulture)

        Dim connu As Bitmap = Nothing
        If _cache.TryGetValue(cle, connu) Then Return connu

        Dim image As Bitmap = Charger(produit, cote)
        If image Is Nothing Then image = Dessiner(produit, cote)

        _cache(cle) = image
        Return image
    End Function

    ''' <summary>
    ''' Vrai si le logo officiel du produit est bien déposé et lisible. Sert à dire à
    ''' l'utilisateur où poser le fichier, tant qu'il n'y est pas.
    ''' </summary>
    Public Shared Function EstFourni(produit As ProduitTransfert) As Boolean
        Return FichierDuLogo(produit) IsNot Nothing
    End Function

    ''' <summary>
    ''' Le chemin où déposer le logo du produit. Il est rendu même si le fichier n'existe pas :
    ''' c'est précisément dans ce cas qu'on a besoin de le montrer.
    ''' </summary>
    Public Shared Function CheminAttendu(produit As ProduitTransfert) As String

        If produit Is Nothing Then Return String.Empty

        Dim depose As String = FichierDuLogo(produit)
        If depose IsNot Nothing Then Return depose

        Return Path.Combine(DossierDesLogos(), produit.Code & ".png")
    End Function

#End Region

#Region "Le fichier officiel"

    ''' <summary>Le dossier des logos, à côté de l'exécutable.</summary>
    Private Shared Function DossierDesLogos() As String

        Try
            Return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, DOSSIER)

        Catch ex As Exception
            ' Un domaine sans dossier de base — cas de figure d'hébergement exotique — ne doit
            ' pas empêcher l'écran de s'afficher : l'emblème dessiné prendra le relais.
            Return String.Empty
        End Try
    End Function

    ''' <summary>Le premier fichier de logo trouvé pour ce produit, ou Nothing.</summary>
    Private Shared Function FichierDuLogo(produit As ProduitTransfert) As String

        If produit Is Nothing Then Return Nothing

        Dim dossier As String = DossierDesLogos()
        If String.IsNullOrEmpty(dossier) Then Return Nothing

        Try
            If Not Directory.Exists(dossier) Then Return Nothing

            For Each extension As String In EXTENSIONS

                Dim chemin As String = Path.Combine(dossier, produit.Code & extension)
                If File.Exists(chemin) Then Return chemin
            Next

        Catch ex As Exception
            ' Dossier inaccessible, chemin trop long, droits refusés : le logo n'est pas là,
            ' et c'est tout ce que l'appelant a besoin de savoir.
            Return Nothing
        End Try

        Return Nothing
    End Function

    ''' <summary>
    ''' Lit le fichier du logo et le rend à la taille demandée, ses proportions conservées et
    ''' l'image centrée. Rend Nothing si le fichier n'existe pas ou ne se lit pas.
    '''
    ''' LE FICHIER N'EST PAS VERROUILLÉ. Image.FromFile garde le fichier ouvert tant que
    ''' l'image vit — et elle vit ici toute la session, puisqu'elle est en cache. Le marketing
    ''' n'aurait alors pas pu remplacer le logo sans fermer l'application. Les octets sont donc
    ''' lus d'un coup, et l'image construite sur eux.
    ''' </summary>
    Private Shared Function Charger(produit As ProduitTransfert, cote As Integer) As Bitmap

        Dim chemin As String = FichierDuLogo(produit)
        If chemin Is Nothing Then Return Nothing

        Try
            Dim octets As Byte() = File.ReadAllBytes(chemin)

            Using flux As New MemoryStream(octets)
                Using source As Image = Image.FromStream(flux)

                    If source.Width <= 0 OrElse source.Height <= 0 Then Return Nothing

                    Dim facteur As Single = Math.Min(CSng(cote) / CSng(source.Width),
                                                     CSng(cote) / CSng(source.Height))

                    Dim large As Integer = Math.Max(1, CInt(Math.Round(source.Width * facteur)))
                    Dim haut As Integer = Math.Max(1, CInt(Math.Round(source.Height * facteur)))

                    Dim rendu As New Bitmap(cote, cote, System.Drawing.Imaging.PixelFormat.Format32bppArgb)

                    Using g As Graphics = Graphics.FromImage(rendu)

                        g.InterpolationMode = InterpolationMode.HighQualityBicubic
                        g.PixelOffsetMode = PixelOffsetMode.HighQuality
                        g.SmoothingMode = SmoothingMode.AntiAlias

                        g.DrawImage(source, New Rectangle((cote - large) \ 2, (cote - haut) \ 2,
                                                          large, haut))
                    End Using

                    Return rendu
                End Using
            End Using

        Catch ex As Exception
            ' Fichier corrompu, format non reconnu, fichier en cours d'écriture : l'emblème
            ' dessiné prend le relais. Un logo illisible ne doit pas empêcher de choisir un
            ' produit et de travailler.
            Return Nothing
        End Try
    End Function

#End Region

#Region "L'emblème de repli"

    ''' <summary>
    ''' Dessine un carré arrondi aux couleurs du produit, portant ses initiales.
    '''
    ''' Volontairement sobre, et volontairement PAS ressemblant : c'est une pastille de
    ''' remplacement, pas une imitation de la marque.
    ''' </summary>
    Private Shared Function Dessiner(produit As ProduitTransfert, cote As Integer) As Bitmap

        Dim rendu As New Bitmap(cote, cote, System.Drawing.Imaging.PixelFormat.Format32bppArgb)

        Using g As Graphics = Graphics.FromImage(rendu)

            g.SmoothingMode = SmoothingMode.AntiAlias
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias

            ' Une marge d'un seizième : le carré ne touche pas le bord, et l'anticrénelage a
            ' la place de travailler sans être coupé.
            Dim marge As Single = Math.Max(1.0F, cote / 16.0F)
            Dim cadre As New RectangleF(marge, marge, cote - 2.0F * marge, cote - 2.0F * marge)
            Dim rayon As Single = cadre.Width / 5.0F

            Using contour As GraphicsPath = CarreArrondi(cadre, rayon)

                Using fond As New SolidBrush(produit.Couleur)
                    g.FillPath(fond, contour)
                End Using

                ' Un liseré sombre translucide : sur un fond clair — l'ambre, par exemple — le
                ' carré se détacherait mal du blanc d'une liste.
                Using plume As New Pen(Color.FromArgb(40, 0, 0, 0), Math.Max(1.0F, cote / 64.0F))
                    g.DrawPath(plume, contour)
                End Using
            End Using

            ' La taille de police suit le nombre d'initiales : « WU » tient large, « RIA » non.
            Dim initiales As String = produit.Initiales
            Dim facteur As Single = If(initiales.Length >= 3, 0.30F, 0.40F)

            Using police As New Font("Segoe UI", cadre.Width * facteur, FontStyle.Bold,
                                     GraphicsUnit.Pixel)
                Using alignement As New StringFormat()

                    alignement.Alignment = StringAlignment.Center
                    alignement.LineAlignment = StringAlignment.Center

                    Using encre As New SolidBrush(produit.CouleurDuTexte)
                        g.DrawString(initiales, police, encre, cadre, alignement)
                    End Using
                End Using
            End Using
        End Using

        Return rendu
    End Function

    ''' <summary>Le contour d'un carré à coins arrondis.</summary>
    Private Shared Function CarreArrondi(cadre As RectangleF, rayon As Single) As GraphicsPath

        Dim chemin As New GraphicsPath()
        Dim diametre As Single = rayon * 2.0F

        chemin.AddArc(cadre.Left, cadre.Top, diametre, diametre, 180.0F, 90.0F)
        chemin.AddArc(cadre.Right - diametre, cadre.Top, diametre, diametre, 270.0F, 90.0F)
        chemin.AddArc(cadre.Right - diametre, cadre.Bottom - diametre, diametre, diametre, 0.0F, 90.0F)
        chemin.AddArc(cadre.Left, cadre.Bottom - diametre, diametre, diametre, 90.0F, 90.0F)
        chemin.CloseFigure()

        Return chemin
    End Function

#End Region

End Class
