Option Strict On
Option Explicit On

Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Globalization
Imports System.Windows.Forms

''' <summary>
''' Les icônes des menus et de la barre d'outils, dessinées par le programme.
'''
''' POURQUOI LES DESSINER PLUTÔT QUE LES IMPORTER
'''
''' Une application qui porte des images doit les transporter : un dossier de PNG à côté de
''' l'exécutable se perd au premier déploiement, et des images incorporées gonflent le projet
''' de fichiers binaires que personne ne peut relire ni corriger. Ici chaque icône tient en
''' quelques traits de GDI+ : le dessin est du texte, il se relit, il se modifie, et il suit
''' l'exécutable sans rien emporter.
'''
''' Second avantage, décisif sur les postes de la banque : le dessin est calculé À LA TAILLE
''' DEMANDÉE. Un poste en 125 % ou 150 % — courant sous Windows 10 et Windows 11 — réclame des
''' images de 20 ou 24 pixels ; une image de 16 pixels agrandie par Windows devient floue,
''' celle-ci reste nette parce qu'elle est retracée.
'''
''' LA GRILLE DE DESSIN
'''
''' Toutes les icônes sont composées dans un carré de 16 unités, quelle que soit la taille
''' finale : la mise à l'échelle est posée une fois sur le contexte graphique (ScaleTransform),
''' et chaque tracé travaille en unités de la grille. Cela évite de recalculer des coordonnées
''' pour chaque taille, et garantit que les dix-neuf dessins restent cohérents entre eux.
'''
''' LES COULEURS
'''
''' Elles sont relevées sur l'icône de l'application fournie par la banque : l'or, le vert et
''' le bleu des cubes. Les icônes de menu appartiennent ainsi visiblement à la même famille que
''' l'icône du bureau et de la barre des tâches. Le contour est une ardoise sombre, choisie
''' plutôt que le noir : à 16 pixels, un contour noir écrase la couleur qu'il entoure.
'''
''' LE GRISÉ DES ENTRÉES DÉSACTIVÉES N'EST PAS ICI
'''
''' Windows Forms grise lui-même l'image d'un élément de menu ou de barre d'outils désactivé,
''' par son moteur de rendu. Produire ici une seconde version grise doublerait le cache pour un
''' résultat que le système fait déjà, et moins bien.
''' </summary>
Public NotInheritable Class IconesWU

    Private Sub New()
    End Sub

    ''' <summary>Côté de la grille de composition. Tous les tracés s'y réfèrent.</summary>
    Private Const COTE_DESSIN As Single = 16.0F

    ''' <summary>
    ''' Nom de la ressource incorporée qui porte l'icône de l'application.
    '''
    ''' Il est composé du RootNamespace du projet et du chemin du fichier, les séparateurs
    ''' devenant des points : c'est la règle de nommage de MSBuild pour une EmbeddedResource.
    ''' Déplacer ou renommer le fichier .ico sans corriger cette constante ferait échouer la
    ''' lecture — silencieusement, l'application gardant alors l'icône par défaut de Windows.
    ''' </summary>
    Private Const RESSOURCE_ICONE As String = "WincompenseTCHAD.Ressources.Wincompense.ico"

    ' La palette, relevée sur l'icône fournie par la banque.
    Private Shared ReadOnly TEINTE_OR As Color = Color.FromArgb(255, 213, 76)
    Private Shared ReadOnly TEINTE_OR_SOMBRE As Color = Color.FromArgb(240, 171, 40)
    Private Shared ReadOnly TEINTE_VERT As Color = Color.FromArgb(121, 187, 0)
    Private Shared ReadOnly TEINTE_BLEU As Color = Color.FromArgb(60, 145, 189)
    Private Shared ReadOnly TEINTE_BLEU_CLAIR As Color = Color.FromArgb(79, 175, 230)
    Private Shared ReadOnly TEINTE_ARDOISE As Color = Color.FromArgb(55, 71, 79)
    Private Shared ReadOnly TEINTE_PAPIER As Color = Color.FromArgb(252, 252, 252)
    Private Shared ReadOnly TEINTE_ALERTE As Color = Color.FromArgb(198, 40, 40)

    ''' <summary>
    ''' Les dessins déjà produits, rangés par icône et par taille.
    '''
    ''' Le cache n'est jamais vidé et ses images ne sont jamais libérées : c'est voulu. Une
    ''' vingtaine d'images de 16 à 24 pixels pèsent quelques dizaines de kilo-octets, elles
    ''' vivent le temps de l'application, et elles sont posées sur des menus qui les gardent.
    ''' Libérer une image encore affichée provoquerait une exception au premier repaint.
    ''' </summary>
    Private Shared ReadOnly _cache As New Dictionary(Of String, Bitmap)()

    Private Shared _iconeApplication As Icon
    Private Shared _iconeCherchee As Boolean

#Region "Accès"

    ''' <summary>
    ''' Rend l'image d'une icône, à la taille demandée, en la dessinant au premier appel.
    ''' </summary>
    ''' <param name="icone">Le dessin voulu.</param>
    ''' <param name="taille">Côté de l'image en pixels ; 16 convient aux menus, 24 à la barre d'outils.</param>
    Public Shared Function Obtenir(icone As IconeWU, Optional taille As Integer = 16) As Bitmap

        ' Une taille aberrante ne doit pas faire tomber l'écran qui la demande : elle est
        ' ramenée dans des bornes raisonnables, et l'icône sort quand même.
        Dim cote As Integer = taille
        If cote < 8 Then cote = 8
        If cote > 256 Then cote = 256

        Dim cle As String = CInt(icone).ToString(CultureInfo.InvariantCulture) &
                            "|" & cote.ToString(CultureInfo.InvariantCulture)

        Dim connue As Bitmap = Nothing
        If _cache.TryGetValue(cle, connue) Then Return connue

        Dim produite As Bitmap = Fabriquer(icone, cote)
        _cache(cle) = produite
        Return produite
    End Function

    ''' <summary>
    ''' Rend l'icône de l'application, lue une fois dans les ressources incorporées.
    '''
    ''' Renvoie Nothing si la ressource est absente : l'application garde alors l'icône par
    ''' défaut de Windows Forms. Une icône manquante est un défaut d'habillage, jamais une
    ''' raison d'empêcher un agent de travailler.
    ''' </summary>
    Public Shared Function IconeApplication() As Icon

        If _iconeCherchee Then Return _iconeApplication
        _iconeCherchee = True

        Try
            Dim assemblage As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly()

            Using flux As System.IO.Stream = assemblage.GetManifestResourceStream(RESSOURCE_ICONE)
                If flux IsNot Nothing Then _iconeApplication = New Icon(flux)
            End Using

        Catch ex As ArgumentException
            ' Ressource présente mais illisible : on s'en passe.
            _iconeApplication = Nothing

        Catch ex As System.IO.IOException
            _iconeApplication = Nothing
        End Try

        Return _iconeApplication
    End Function

    ''' <summary>
    ''' Pose l'icône de l'application sur une fenêtre, si elle a pu être lue.
    '''
    ''' Écrire Nothing dans Form.Icon ne laisserait pas la fenêtre en l'état : cela rétablirait
    ''' l'icône par défaut. Le test n'est donc pas une précaution, il porte le comportement.
    ''' </summary>
    Public Shared Sub Habiller(ecran As Form)

        If ecran Is Nothing Then Return

        Dim marque As Icon = IconeApplication()
        If marque Is Nothing Then Return

        ecran.Icon = marque
    End Sub

    ''' <summary>
    ''' Rend une copie de l'image, marquée d'une pastille portant un nombre.
    '''
    ''' Elle sert au bouton des autorisations en attente : dans la barre d'outils, le bouton
    ''' n'affiche pas de texte, et le nombre écrit dans le menu ne s'y verrait pas. Au-delà de
    ''' 99, la pastille affiche « 99+ » : trois caractères sont déjà la limite de ce qui reste
    ''' lisible à 24 pixels, et le compte exact se lit dans le menu.
    '''
    ''' L'image rendue est neuve à chaque appel — elle n'est pas mise en cache, puisque le
    ''' nombre change — et il revient à l'appelant de libérer la précédente.
    ''' </summary>
    ''' <param name="source">Image de base ; elle n'est pas modifiée.</param>
    ''' <param name="nombre">Nombre à inscrire ; zéro ou moins rend la source telle quelle.</param>
    Public Shared Function AvecPastille(source As Bitmap, nombre As Integer) As Bitmap

        If source Is Nothing Then Return Nothing
        If nombre <= 0 Then Return source

        Dim marquee As New Bitmap(source.Width, source.Height,
                                  System.Drawing.Imaging.PixelFormat.Format32bppArgb)

        Using g As Graphics = Graphics.FromImage(marquee)

            g.SmoothingMode = SmoothingMode.AntiAlias
            g.InterpolationMode = InterpolationMode.HighQualityBicubic
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias
            g.DrawImage(source, 0, 0, source.Width, source.Height)

            Dim diametre As Single = source.Width * 0.66F
            Dim gauche As Single = source.Width - diametre
            Dim ecrit As String = If(nombre > 99, "99+",
                                     nombre.ToString(CultureInfo.InvariantCulture))

            Using fond As New SolidBrush(TEINTE_ALERTE)
                g.FillEllipse(fond, gauche, 0.0F, diametre, diametre)
            End Using

            ' Un liseré clair détache la pastille du dessin qu'elle recouvre : sans lui, une
            ' pastille rouge posée sur un aplat sombre disparaît.
            Using lisere As New Pen(Color.White, Math.Max(1.0F, source.Width / 16.0F))
                g.DrawEllipse(lisere, gauche, 0.0F, diametre, diametre)
            End Using

            Dim corps As Single = diametre * If(ecrit.Length > 2, 0.42F, 0.58F)

            Using police As New Font("Segoe UI", corps, FontStyle.Bold, GraphicsUnit.Pixel)
                Using encre As New SolidBrush(Color.White)
                    Using mise As New StringFormat()
                        mise.Alignment = StringAlignment.Center
                        mise.LineAlignment = StringAlignment.Center
                        g.DrawString(ecrit, police, encre,
                                     New RectangleF(gauche, 0.0F, diametre, diametre), mise)
                    End Using
                End Using
            End Using
        End Using

        Return marquee
    End Function

#End Region

#Region "Fabrication"

    ''' <summary>
    ''' Dessine une icône à la taille demandée, dans la grille de 16 unités.
    ''' </summary>
    Private Shared Function Fabriquer(icone As IconeWU, cote As Integer) As Bitmap

        Dim planche As New Bitmap(cote, cote,
                                  System.Drawing.Imaging.PixelFormat.Format32bppArgb)

        Using g As Graphics = Graphics.FromImage(planche)

            g.SmoothingMode = SmoothingMode.AntiAlias
            g.PixelOffsetMode = PixelOffsetMode.HighQuality
            g.InterpolationMode = InterpolationMode.HighQualityBicubic

            Dim facteur As Single = CSng(cote) / COTE_DESSIN
            g.ScaleTransform(facteur, facteur)

            Tracer(g, icone)
        End Using

        Return planche
    End Function

    ''' <summary>
    ''' Aiguille vers le tracé de chaque icône. Une icône inconnue ne dessine rien : mieux vaut
    ''' un menu sans image qu'une exception au premier affichage.
    ''' </summary>
    Private Shared Sub Tracer(g As Graphics, icone As IconeWU)

        Select Case icone
            Case IconeWU.Balance : TracerBalance(g)
            Case IconeWU.Graphique : TracerGraphique(g)
            Case IconeWU.Batiment : TracerBatiment(g)
            Case IconeWU.Archive : TracerArchive(g)
            Case IconeWU.Piece : TracerPiece(g)
            Case IconeWU.Silhouette : TracerSilhouette(g)
            Case IconeWU.Groupe : TracerGroupe(g)
            Case IconeWU.Coche : TracerCoche(g)
            Case IconeWU.Registre : TracerRegistre(g)
            Case IconeWU.Curseurs : TracerCurseurs(g)
            Case IconeWU.Dossier : TracerDossier(g)
            Case IconeWU.Cle : TracerCle(g)
            Case IconeWU.Utilisateurs : TracerUtilisateurs(g)
            Case IconeWU.BaseDeDonnees : TracerBaseDeDonnees(g)
            Case IconeWU.Cascade : TracerCascade(g)
            Case IconeWU.MosaiqueH : TracerMosaiqueH(g)
            Case IconeWU.MosaiqueV : TracerMosaiqueV(g)
            Case IconeWU.FermerTout : TracerFermerTout(g)
            Case IconeWU.Sortie : TracerSortie(g)
        End Select
    End Sub

#End Region

#Region "Outils de tracé"

    ''' <summary>
    ''' Un stylo de contour : ardoise, bouts et angles arrondis.
    '''
    ''' Les bouts arrondis ne sont pas une coquetterie : à 16 pixels, un trait à bout carré
    ''' déborde d'un demi-pixel et fait mordre les angles.
    ''' </summary>
    Private Shared Function Contour(Optional epaisseur As Single = 1.1F) As Pen

        Dim stylo As New Pen(TEINTE_ARDOISE, epaisseur)
        stylo.LineJoin = LineJoin.Round
        stylo.StartCap = LineCap.Round
        stylo.EndCap = LineCap.Round
        Return stylo
    End Function

    ''' <summary>Remplit puis cerne un rectangle, en une fois.</summary>
    Private Shared Sub Boite(g As Graphics, x As Single, y As Single, l As Single, h As Single,
                             teinte As Color, Optional cerne As Boolean = True)

        Using fond As New SolidBrush(teinte)
            g.FillRectangle(fond, x, y, l, h)
        End Using

        If Not cerne Then Return

        Using stylo As Pen = Contour()
            g.DrawRectangle(stylo, x, y, l, h)
        End Using
    End Sub

    ''' <summary>Remplit puis cerne un disque, désigné par son centre et son rayon.</summary>
    Private Shared Sub Disque(g As Graphics, cx As Single, cy As Single, rayon As Single,
                              teinte As Color, Optional cerne As Boolean = True)

        Using fond As New SolidBrush(teinte)
            g.FillEllipse(fond, cx - rayon, cy - rayon, rayon * 2.0F, rayon * 2.0F)
        End Using

        If Not cerne Then Return

        Using stylo As Pen = Contour()
            g.DrawEllipse(stylo, cx - rayon, cy - rayon, rayon * 2.0F, rayon * 2.0F)
        End Using
    End Sub

    ''' <summary>Remplit puis cerne un polygone fermé.</summary>
    Private Shared Sub Forme(g As Graphics, sommets As PointF(), teinte As Color,
                             Optional cerne As Boolean = True)

        Using fond As New SolidBrush(teinte)
            g.FillPolygon(fond, sommets)
        End Using

        If Not cerne Then Return

        Using stylo As Pen = Contour()
            g.DrawPolygon(stylo, sommets)
        End Using
    End Sub

    ''' <summary>
    ''' Dessine une petite fenêtre : corps clair, bandeau de titre coloré, contour.
    '''
    ''' Trois icônes de disposition — cascade et mosaïques — et celle de fermeture reposent sur
    ''' ce motif ; l'écrire une fois les garde identiques entre elles.
    ''' </summary>
    Private Shared Sub Fenetre(g As Graphics, x As Single, y As Single, l As Single, h As Single,
                               bandeau As Color)

        Using fond As New SolidBrush(TEINTE_PAPIER)
            g.FillRectangle(fond, x, y, l, h)
        End Using

        Using titre As New SolidBrush(bandeau)
            g.FillRectangle(titre, x, y, l, Math.Min(h, 2.2F))
        End Using

        Using stylo As Pen = Contour()
            g.DrawRectangle(stylo, x, y, l, h)
        End Using
    End Sub

#End Region

#Region "Les dix-neuf dessins"

    ''' <summary>Balance à deux plateaux : le traitement de la compense, qui équilibre la journée.</summary>
    Private Shared Sub TracerBalance(g As Graphics)

        ' Le fléau et le mât sont tracés plus épais que le contour ordinaire : à 16 pixels,
        ' un trait de 1,1 unité s'efface derrière les plateaux qui l'encadrent.
        Using stylo As Pen = Contour(1.5F)
            g.DrawLine(stylo, 2.5F, 4.6F, 13.5F, 4.6F)          ' le fléau
            g.DrawLine(stylo, 8.0F, 4.6F, 8.0F, 12.2F)          ' le mât
        End Using

        Using suspentes As Pen = Contour(1.1F)
            g.DrawLine(suspentes, 2.5F, 4.6F, 2.5F, 6.5F)
            g.DrawLine(suspentes, 13.5F, 4.6F, 13.5F, 6.5F)
        End Using

        Forme(g, New PointF() {New PointF(0.6F, 6.5F), New PointF(4.4F, 6.5F),
                               New PointF(3.5F, 9.1F), New PointF(1.5F, 9.1F)}, TEINTE_BLEU)

        Forme(g, New PointF() {New PointF(11.6F, 6.5F), New PointF(15.4F, 6.5F),
                               New PointF(14.5F, 9.1F), New PointF(12.5F, 9.1F)}, TEINTE_VERT)

        Forme(g, New PointF() {New PointF(5.4F, 13.6F), New PointF(10.6F, 13.6F),
                               New PointF(9.6F, 12.2F), New PointF(6.4F, 12.2F)}, TEINTE_OR)

        Disque(g, 8.0F, 3.4F, 1.2F, TEINTE_OR)
    End Sub

    ''' <summary>Trois barres montantes : le rapport d'activité des sous-agents.</summary>
    Private Shared Sub TracerGraphique(g As Graphics)

        Boite(g, 2.0F, 8.2F, 3.0F, 5.0F, TEINTE_VERT)
        Boite(g, 6.5F, 5.2F, 3.0F, 8.0F, TEINTE_OR)
        Boite(g, 11.0F, 2.4F, 3.0F, 10.8F, TEINTE_BLEU)

        Using stylo As Pen = Contour(1.2F)
            g.DrawLine(stylo, 1.2F, 13.8F, 14.8F, 13.8F)
        End Using
    End Sub

    ''' <summary>Immeuble à fenêtres : les agences propres de la banque.</summary>
    Private Shared Sub TracerBatiment(g As Graphics)

        Boite(g, 3.2F, 3.2F, 9.6F, 10.4F, TEINTE_OR)

        Using vitres As New SolidBrush(TEINTE_BLEU)
            For Each x As Single In New Single() {4.7F, 7.2F, 9.7F}
                g.FillRectangle(vitres, x, 4.8F, 1.6F, 1.6F)
                g.FillRectangle(vitres, x, 7.3F, 1.6F, 1.6F)
            Next
        End Using

        Using porte As New SolidBrush(TEINTE_ARDOISE)
            g.FillRectangle(porte, 7.0F, 10.2F, 2.2F, 3.4F)
        End Using
    End Sub

    ''' <summary>Boîte d'archives à couvercle : les pièces comptables conservées.</summary>
    Private Shared Sub TracerArchive(g As Graphics)

        Boite(g, 2.4F, 6.4F, 11.2F, 7.2F, TEINTE_OR)
        Boite(g, 1.4F, 3.2F, 13.2F, 3.2F, TEINTE_OR_SOMBRE)

        Using poignee As New SolidBrush(TEINTE_PAPIER)
            g.FillRectangle(poignee, 6.4F, 8.4F, 3.2F, 1.6F)
        End Using

        Using stylo As Pen = Contour(0.9F)
            g.DrawRectangle(stylo, 6.4F, 8.4F, 3.2F, 1.6F)
        End Using
    End Sub

    ''' <summary>Deux pièces de monnaie : les commissions encaissées par la banque.</summary>
    Private Shared Sub TracerPiece(g As Graphics)

        ' Les deux disques sont nettement décalés : trop rapprochés, ils ne forment qu'une
        ' tache à 16 pixels, et l'icône cesse de dire « des pièces » pour dire « un rond ».
        Disque(g, 5.0F, 10.6F, 3.8F, TEINTE_OR_SOMBRE)
        Disque(g, 10.2F, 5.6F, 4.4F, TEINTE_OR)

        Using stylo As Pen = Contour(0.9F)
            g.DrawEllipse(stylo, 7.9F, 3.3F, 4.6F, 4.6F)
        End Using
    End Sub

    ''' <summary>Une silhouette : les sous-agents.</summary>
    Private Shared Sub TracerSilhouette(g As Graphics)

        Disque(g, 8.0F, 5.0F, 2.8F, TEINTE_OR)

        Using buste As New SolidBrush(TEINTE_OR)
            g.FillPie(buste, 2.2F, 8.6F, 11.6F, 11.0F, 180.0F, 180.0F)
        End Using

        Using stylo As Pen = Contour()
            g.DrawArc(stylo, 2.2F, 8.6F, 11.6F, 11.0F, 180.0F, 180.0F)
        End Using
    End Sub

    ''' <summary>Trois pastilles reliées : les groupes statistiques.</summary>
    Private Shared Sub TracerGroupe(g As Graphics)

        Using liens As Pen = Contour(1.0F)
            g.DrawLine(liens, 4.4F, 5.2F, 11.6F, 5.2F)
            g.DrawLine(liens, 4.4F, 5.2F, 8.0F, 11.0F)
            g.DrawLine(liens, 11.6F, 5.2F, 8.0F, 11.0F)
        End Using

        Disque(g, 4.4F, 5.2F, 2.3F, TEINTE_VERT)
        Disque(g, 11.6F, 5.2F, 2.3F, TEINTE_BLEU)
        Disque(g, 8.0F, 11.0F, 2.5F, TEINTE_OR)
    End Sub

    ''' <summary>Une coche dans une pastille : les autorisations du référentiel.</summary>
    Private Shared Sub TracerCoche(g As Graphics)

        Disque(g, 8.0F, 8.0F, 6.2F, TEINTE_VERT)

        Using marque As New Pen(Color.White, 2.0F)
            marque.StartCap = LineCap.Round
            marque.EndCap = LineCap.Round
            marque.LineJoin = LineJoin.Round
            g.DrawLines(marque, New PointF() {New PointF(4.8F, 8.2F),
                                              New PointF(7.1F, 10.6F),
                                              New PointF(11.4F, 5.4F)})
        End Using
    End Sub

    ''' <summary>Un registre ouvert : les comptes systèmes, le plan comptable de l'application.</summary>
    Private Shared Sub TracerRegistre(g As Graphics)

        Boite(g, 2.4F, 2.4F, 11.2F, 11.4F, TEINTE_PAPIER)

        Using tranche As New SolidBrush(TEINTE_BLEU)
            g.FillRectangle(tranche, 2.4F, 2.4F, 2.2F, 11.4F)
        End Using

        Using stylo As Pen = Contour(0.9F)
            g.DrawRectangle(stylo, 2.4F, 2.4F, 11.2F, 11.4F)

            For Each y As Single In New Single() {5.0F, 7.2F, 9.4F, 11.6F}
                g.DrawLine(stylo, 5.8F, y, 12.0F, y)
            Next
        End Using
    End Sub

    ''' <summary>Deux curseurs de réglage : les options de traitement.</summary>
    Private Shared Sub TracerCurseurs(g As Graphics)

        Using rail As Pen = Contour(1.3F)
            g.DrawLine(rail, 2.0F, 5.4F, 14.0F, 5.4F)
            g.DrawLine(rail, 2.0F, 10.6F, 14.0F, 10.6F)
        End Using

        Disque(g, 5.4F, 5.4F, 1.9F, TEINTE_OR)
        Disque(g, 10.6F, 10.6F, 1.9F, TEINTE_BLEU)
    End Sub

    ''' <summary>Une chemise : le fichier de secours du paramétrage.</summary>
    Private Shared Sub TracerDossier(g As Graphics)

        Forme(g, New PointF() {New PointF(1.6F, 3.6F), New PointF(6.2F, 3.6F),
                               New PointF(7.4F, 5.4F), New PointF(14.4F, 5.4F),
                               New PointF(14.4F, 13.2F), New PointF(1.6F, 13.2F)}, TEINTE_OR_SOMBRE)

        Forme(g, New PointF() {New PointF(1.6F, 7.0F), New PointF(14.4F, 7.0F),
                               New PointF(14.4F, 13.2F), New PointF(1.6F, 13.2F)}, TEINTE_OR)
    End Sub

    ''' <summary>Une clé : le mot de passe de l'utilisateur connecté.</summary>
    Private Shared Sub TracerCle(g As Graphics)

        Boite(g, 6.8F, 7.1F, 7.6F, 1.9F, TEINTE_OR)

        Using dents As New SolidBrush(TEINTE_OR)
            g.FillRectangle(dents, 10.6F, 9.0F, 1.1F, 1.9F)
            g.FillRectangle(dents, 12.8F, 9.0F, 1.1F, 1.9F)
        End Using

        Using stylo As Pen = Contour(0.9F)
            g.DrawRectangle(stylo, 10.6F, 9.0F, 1.1F, 1.9F)
            g.DrawRectangle(stylo, 12.8F, 9.0F, 1.1F, 1.9F)
        End Using

        Disque(g, 4.6F, 8.0F, 3.4F, TEINTE_BLEU)
        Disque(g, 4.6F, 8.0F, 1.3F, TEINTE_PAPIER)
    End Sub

    ''' <summary>Deux silhouettes : les utilisateurs et leurs connexions.</summary>
    Private Shared Sub TracerUtilisateurs(g As Graphics)

        ' La silhouette du fond est tracée d'abord : celle du premier plan la recouvre en
        ' partie, et c'est ce recouvrement qui fait lire « deux personnes » plutôt que « deux
        ' ronds ». Les deux bustes portent des noms distincts pour que l'ordre reste lisible.
        Disque(g, 11.0F, 5.2F, 2.1F, TEINTE_BLEU)

        Using busteArriere As New SolidBrush(TEINTE_BLEU)
            g.FillPie(busteArriere, 7.4F, 8.0F, 7.2F, 7.0F, 180.0F, 180.0F)
        End Using

        Using styloArriere As Pen = Contour()
            g.DrawArc(styloArriere, 7.4F, 8.0F, 7.2F, 7.0F, 180.0F, 180.0F)
        End Using

        Disque(g, 5.8F, 5.6F, 2.5F, TEINTE_OR)

        Using busteAvant As New SolidBrush(TEINTE_OR)
            g.FillPie(busteAvant, 1.4F, 8.8F, 8.8F, 8.6F, 180.0F, 180.0F)
        End Using

        Using styloAvant As Pen = Contour()
            g.DrawArc(styloAvant, 1.4F, 8.8F, 8.8F, 8.6F, 180.0F, 180.0F)
        End Using
    End Sub

    ''' <summary>Un cylindre à bandes : la connexion à la base de données.</summary>
    Private Shared Sub TracerBaseDeDonnees(g As Graphics)

        Using corps As New SolidBrush(TEINTE_BLEU)
            g.FillRectangle(corps, 2.6F, 3.8F, 10.8F, 8.4F)
            g.FillEllipse(corps, 2.6F, 10.4F, 10.8F, 3.6F)
        End Using

        Using couvercle As New SolidBrush(TEINTE_BLEU_CLAIR)
            g.FillEllipse(couvercle, 2.6F, 2.0F, 10.8F, 3.6F)
        End Using

        Using stylo As Pen = Contour()
            g.DrawArc(stylo, 2.6F, 10.4F, 10.8F, 3.6F, 0.0F, 180.0F)
            g.DrawLine(stylo, 2.6F, 3.8F, 2.6F, 12.2F)
            g.DrawLine(stylo, 13.4F, 3.8F, 13.4F, 12.2F)
            g.DrawEllipse(stylo, 2.6F, 2.0F, 10.8F, 3.6F)
            g.DrawArc(stylo, 2.6F, 5.0F, 10.8F, 3.6F, 0.0F, 180.0F)
            g.DrawArc(stylo, 2.6F, 7.7F, 10.8F, 3.6F, 0.0F, 180.0F)
        End Using
    End Sub

    ''' <summary>Trois fenêtres décalées : la disposition en cascade.</summary>
    Private Shared Sub TracerCascade(g As Graphics)

        Fenetre(g, 1.4F, 1.8F, 8.4F, 6.4F, TEINTE_BLEU_CLAIR)
        Fenetre(g, 3.8F, 4.6F, 8.4F, 6.4F, TEINTE_BLEU)
        Fenetre(g, 6.2F, 7.4F, 8.4F, 6.4F, TEINTE_OR)
    End Sub

    ''' <summary>Deux fenêtres empilées : la mosaïque horizontale, qui partage la hauteur.</summary>
    Private Shared Sub TracerMosaiqueH(g As Graphics)

        Fenetre(g, 2.0F, 2.4F, 12.0F, 4.8F, TEINTE_BLEU)
        Fenetre(g, 2.0F, 8.8F, 12.0F, 4.8F, TEINTE_OR)
    End Sub

    ''' <summary>Deux fenêtres côte à côte : la mosaïque verticale, qui partage la largeur.</summary>
    Private Shared Sub TracerMosaiqueV(g As Graphics)

        Fenetre(g, 2.0F, 2.4F, 5.2F, 11.2F, TEINTE_BLEU)
        Fenetre(g, 8.8F, 2.4F, 5.2F, 11.2F, TEINTE_OR)
    End Sub

    ''' <summary>Une fenêtre barrée : la fermeture de toutes les fenêtres.</summary>
    Private Shared Sub TracerFermerTout(g As Graphics)

        Fenetre(g, 2.0F, 2.6F, 12.0F, 10.8F, TEINTE_BLEU)

        Using croix As New Pen(TEINTE_ALERTE, 1.9F)
            croix.StartCap = LineCap.Round
            croix.EndCap = LineCap.Round
            g.DrawLine(croix, 5.6F, 7.2F, 10.4F, 12.0F)
            g.DrawLine(croix, 10.4F, 7.2F, 5.6F, 12.0F)
        End Using
    End Sub

    ''' <summary>Une porte et une flèche : quitter l'application.</summary>
    Private Shared Sub TracerSortie(g As Graphics)

        Boite(g, 1.6F, 2.0F, 6.4F, 12.0F, TEINTE_OR)

        Using bouton As New SolidBrush(TEINTE_ARDOISE)
            g.FillEllipse(bouton, 6.0F, 7.4F, 1.3F, 1.3F)
        End Using

        Using fleche As New Pen(TEINTE_ARDOISE, 1.6F)
            fleche.StartCap = LineCap.Round
            fleche.EndCap = LineCap.Round
            g.DrawLine(fleche, 9.2F, 8.0F, 13.6F, 8.0F)
        End Using

        Forme(g, New PointF() {New PointF(11.8F, 5.2F), New PointF(15.0F, 8.0F),
                               New PointF(11.8F, 10.8F)}, TEINTE_ARDOISE, False)
    End Sub

#End Region

End Class

''' <summary>
''' Les icônes disponibles. Chaque valeur correspond à un tracé et à un seul.
'''
''' L'énumération est nommée par ce que l'icône REPRÉSENTE — une balance, un registre — et non
''' par l'écran qui l'emploie. Le jour où le rapport des agences propres changera de nom, son
''' icône n'aura pas à changer avec lui, et une même image peut servir à deux endroits sans
''' porter le nom de l'un des deux.
''' </summary>
Public Enum IconeWU
    Balance = 1
    Graphique = 2
    Batiment = 3
    Archive = 4
    Piece = 5
    Silhouette = 6
    Groupe = 7
    Coche = 8
    Registre = 9
    Curseurs = 10
    Dossier = 11
    Cle = 12
    Utilisateurs = 13
    BaseDeDonnees = 14
    Cascade = 15
    MosaiqueH = 16
    MosaiqueV = 17
    FermerTout = 18
    Sortie = 19
End Enum
