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
    ''' Nom du fichier de l'icône, cherché par la FIN du nom de ressource.
    '''
    ''' POURQUOI PAS UN NOM COMPLET ÉCRIT EN DUR. Il l'était, et c'était un piège : le nom
    ''' d'une ressource incorporée n'est pas composé de la même façon en VB.NET et en C#.
    ''' En C#, MSBuild écrit « RootNamespace.Dossier.Fichier.ico » ; **en VB.NET il ignore le
    ''' dossier** et écrit « RootNamespace.Fichier.ico ». Un nom complet écrit à la main est
    ''' donc faux une fois sur deux, et il échoue EN SILENCE : la lecture rend Nothing, et
    ''' l'application garde l'icône par défaut de Windows sans rien signaler.
    '''
    ''' Chercher par la fin du nom fait tomber les trois écritures possibles — celle de VB,
    ''' celle de C#, et le LogicalName posé dans le .vbproj — sans dépendre d'aucune.
    ''' </summary>
    Private Const NOM_FICHIER_ICONE As String = "Wincompense.ico"

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
    ''' Rend l'icône de l'application, lue une fois pour toutes.
    '''
    ''' DEUX SOURCES, ET C'EST VOULU. L'icône est d'abord cherchée dans les ressources
    ''' incorporées, qui portent toutes les tailles — 16, 24, 32, 48 et 64 pixels — si bien
    ''' que la barre de titre et la barre des tâches prennent chacune la sienne, sans
    ''' agrandissement. Si elle n'y est pas, elle est relue DANS L'EXÉCUTABLE lui-même, où
    ''' ApplicationIcon l'a gravée : cette seconde lecture ne rend qu'une taille, mais elle
    ''' ne dépend ni du nom de la ressource ni de la façon dont MSBuild l'a composé.
    '''
    ''' Les deux échouent rarement ensemble : la première tient au nom de la ressource, la
    ''' seconde au fichier .exe. C'est exactement pourquoi il y en a deux — la version
    ''' précédente n'avait que la première, son nom était faux, et l'application n'a porté
    ''' aucune icône sans que rien ne le dise.
    '''
    ''' Renvoie Nothing si les deux échouent : l'application garde alors l'icône par défaut
    ''' de Windows Forms. Une icône manquante est un défaut d'habillage, jamais une raison
    ''' d'empêcher un agent de travailler.
    ''' </summary>
    Public Shared Function IconeApplication() As Icon

        If _iconeCherchee Then Return _iconeApplication
        _iconeCherchee = True

        _iconeApplication = LireDansLesRessources()

        If _iconeApplication Is Nothing Then
            _iconeApplication = LireDansLExecutable()
        End If

        Return _iconeApplication
    End Function

    ''' <summary>
    ''' Cherche l'icône parmi les ressources incorporées, par la fin de leur nom.
    ''' </summary>
    Private Shared Function LireDansLesRessources() As Icon

        Try
            Dim assemblage As System.Reflection.Assembly =
                System.Reflection.Assembly.GetExecutingAssembly()

            For Each nomRessource As String In assemblage.GetManifestResourceNames()

                If Not nomRessource.EndsWith(NOM_FICHIER_ICONE,
                                             StringComparison.OrdinalIgnoreCase) Then
                    Continue For
                End If

                Using flux As System.IO.Stream = assemblage.GetManifestResourceStream(nomRessource)
                    If flux IsNot Nothing Then Return New Icon(flux)
                End Using
            Next

        Catch ex As ArgumentException
            ' Ressource trouvée mais illisible : l'exécutable prendra le relais.
            Return Nothing

        Catch ex As System.IO.IOException
            Return Nothing
        End Try

        Return Nothing
    End Function

    ''' <summary>
    ''' Relit l'icône gravée dans l'exécutable par ApplicationIcon.
    '''
    ''' Elle ne rend qu'une taille, que Windows redimensionne au besoin : c'est un recours,
    ''' pas le chemin normal. Mais c'est un recours qui ne peut pas se tromper de nom.
    ''' </summary>
    Private Shared Function LireDansLExecutable() As Icon

        Try
            Return Icon.ExtractAssociatedIcon(Application.ExecutablePath)

        Catch ex As ArgumentException
            ' Chemin introuvable ou exécutable sans icône : il n'y en aura pas.
            Return Nothing

        Catch ex As System.IO.IOException
            Return Nothing
        End Try
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
        If marque IsNot Nothing Then ecran.Icon = marque

        ' L'icône manquante ne doit pas priver les boutons des leurs. Elles ne viennent pas
        ' de la même source — le fichier .ico pour l'une, des tracés pour les autres — et
        ' les faire échouer ensemble reviendrait à punir deux fois la même absence.
        HabillerLesBoutons(ecran)
    End Sub

    ''' <summary>
    ''' Pose son icône sur chaque bouton de la fenêtre, d'après son NOM.
    '''
    ''' POURQUOI PAR LE NOM, ET NON DANS CHAQUE DESIGNER. Trente-huit boutons répartis sur
    ''' vingt-cinq fenêtres, c'est trente-huit endroits à retoucher, et trente-huit endroits
    ''' à ne pas oublier quand un écran s'ajoute. La table dit une fois pour toutes que
    ''' « btnSupprimer » porte une corbeille, et tout bouton qui s'appellera ainsi demain la
    ''' portera sans qu'on y pense.
    '''
    ''' UN BOUTON QUI PORTE DÉJÀ UNE IMAGE N'EST PAS TOUCHÉ : l'habillage automatique ne
    ''' doit jamais défaire un choix explicite fait dans un Designer.
    '''
    ''' UN NOM INCONNU LAISSE LE BOUTON NU, sans erreur. Un écran neuf dont les boutons ne
    ''' sont pas encore dans la table doit s'ouvrir normalement ; verif_icones_boutons.py
    ''' signale l'oubli au développeur, l'utilisateur n'a pas à le découvrir.
    ''' </summary>
    Public Shared Sub HabillerLesBoutons(ecran As Form)

        If ecran Is Nothing Then Return

        ' La taille suit la police de la fenêtre, qui suit elle-même la mise à l'échelle de
        ' Windows : les fenêtres sont en AutoScaleMode.Font, donc sur un poste réglé à 125 %
        ' la police grandit — et l'icône avec elle, RETRACÉE à la bonne taille plutôt
        ' qu'étirée depuis 16 pixels.
        Dim cote As Integer = Math.Max(16, ecran.Font.Height + 2)

        PoserSurLesBoutons(ecran.Controls, cote)
    End Sub

    ''' <summary>
    ''' Parcourt les contrôles ET LEURS ENFANTS : la moitié des boutons du projet vivent
    ''' dans un GroupBox ou un Panel, et un parcours de surface les manquerait tous.
    ''' </summary>
    Private Shared Sub PoserSurLesBoutons(controles As Control.ControlCollection, cote As Integer)

        If controles Is Nothing Then Return

        For Each controle As Control In controles

            Dim bouton As Button = TryCast(controle, Button)
            If bouton IsNot Nothing Then PoserSurUnBouton(bouton, cote)

            If controle.HasChildren Then PoserSurLesBoutons(controle.Controls, cote)
        Next
    End Sub

    Private Shared Sub PoserSurUnBouton(bouton As Button, cote As Integer)

        If bouton.Image IsNot Nothing Then Return

        Dim icone As IconeWU = IconeDuBouton(bouton.Name)
        If icone = IconeWU.Aucune Then Return

        bouton.Image = Obtenir(icone, cote)

        ' ImageBeforeText place l'image et le texte CÔTE À CÔTE, puis centre l'ensemble.
        ' Poser l'image à gauche en laissant le texte centré les ferait se chevaucher sur
        ' les boutons étroits — « Fermer » n'en fait que cent.
        bouton.ImageAlign = ContentAlignment.MiddleCenter
        bouton.TextAlign = ContentAlignment.MiddleCenter
        bouton.TextImageRelation = TextImageRelation.ImageBeforeText
    End Sub

    ''' <summary>
    ''' L'icône d'un bouton, d'après son nom. Aucune si le nom n'est pas connu.
    '''
    ''' La comparaison ignore la casse : Visual Basic ne la distingue pas, et un bouton
    ''' nommé « btnfermer » dans un Designer retouché à la main doit être habillé comme
    ''' les autres.
    ''' </summary>
    Public Shared Function IconeDuBouton(nom As String) As IconeWU

        Dim cherche As String = If(nom, String.Empty).Trim()
        If cherche.Length = 0 Then Return IconeWU.Aucune

        Dim icone As IconeWU
        If _iconesDesBoutons.TryGetValue(cherche, icone) Then Return icone

        Return IconeWU.Aucune
    End Function

    ''' <summary>
    ''' Le nom de chaque bouton du projet, et le dessin qu'il porte.
    '''
    ''' ELLE EST RANGÉE PAR INTENTION, ET NON PAR ORDRE ALPHABÉTIQUE. Ce qui compte en la
    ''' relisant, c'est de voir que « btnValider » porte le même dessin que
    ''' « btnEnregistrer » et « btnRejeter » le même que « btnAnnuler » : deux boutons qui
    ''' font la même chose doivent se ressembler, et l'alphabet les aurait séparés.
    ''' </summary>
    Private Shared ReadOnly _iconesDesBoutons As Dictionary(Of String, IconeWU) =
        ConstruireLaTableDesBoutons()

    Private Shared Function ConstruireLaTableDesBoutons() As Dictionary(Of String, IconeWU)

        Dim table As New Dictionary(Of String, IconeWU)(StringComparer.OrdinalIgnoreCase)

        ' --- Enregistrer, fermer, renoncer -------------------------------------------------
        table("btnEnregistrer") = IconeWU.Enregistrer
        table("btnValider") = IconeWU.Enregistrer
        table("btnFermer") = IconeWU.Fermer
        table("btnRenoncer") = IconeWU.Annuler
        table("btnAnnuler") = IconeWU.Annuler
        table("btnRejeter") = IconeWU.Annuler

        ' --- Le référentiel : créer, modifier, supprimer -------------------------------------
        table("btnNouveau") = IconeWU.Nouveau
        table("btnNouveauGroupe") = IconeWU.Nouveau
        table("btnModifier") = IconeWU.Modifier
        table("btnSupprimer") = IconeWU.Supprimer
        table("btnCopier") = IconeWU.Copier

        ' --- Relire, resynchroniser -----------------------------------------------------------
        table("btnActualiser") = IconeWU.Actualiser
        table("btnActualiserHistorique") = IconeWU.Actualiser
        table("btnActualiserJournal") = IconeWU.Actualiser
        table("btnSynchroniser") = IconeWU.Actualiser

        ' --- Les deux rapports de la journée, et les fichiers en général ---------------------
        ' btnReglement est PROPRE AU TCHAD : la compensation y part de deux rapports, activité
        ' et règlement. L'application centrafricaine en charge deux autres, et sa table ne
        ' connaît donc pas ce nom. C'est la seule entrée que le portage a dû ajouter.
        table("btnActivite") = IconeWU.Telecharger
        table("btnReglement") = IconeWU.Telecharger
        table("btnCharger") = IconeWU.Telecharger
        table("btnChoisir") = IconeWU.Dossier

        ' --- Ce que la journée produit --------------------------------------------------------
        table("btnAfficher") = IconeWU.Calculer
        table("btnGenererPiece") = IconeWU.Piece
        table("btnPiece") = IconeWU.Piece
        table("btnPieceAccount") = IconeWU.Piece
        table("btnPieceGroupe") = IconeWU.Groupe
        table("btnParGroupe") = IconeWU.Groupe
        table("btnBordereau") = IconeWU.Registre
        table("btnCoreBanking") = IconeWU.BaseDeDonnees
        table("btnExporter") = IconeWU.Exporter
        table("btnMois") = IconeWU.Calendrier

        ' --- Le choix du produit de transfert -------------------------------------------------
        table("btnOuvrir") = IconeWU.Coche
        table("btnQuitter") = IconeWU.Sortie
        table("btnCouleur") = IconeWU.Modifier

        ' --- Le double regard : ce qui s'autorise et ce qui se vise --------------------------
        table("btnAutoriser") = IconeWU.Coche
        table("btnViser") = IconeWU.Coche
        table("btnDeposer") = IconeWU.Coche

        ' --- Les accès, et la base -------------------------------------------------------------
        table("btnConnexion") = IconeWU.Cle
        table("btnReinitialiser") = IconeWU.Cle
        table("btnDeverrouiller") = IconeWU.Cle
        table("btnParametres") = IconeWU.BaseDeDonnees
        table("btnTester") = IconeWU.BaseDeDonnees
        table("btnPreparer") = IconeWU.BaseDeDonnees

        Return table
    End Function

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
            Case IconeWU.Pourcentage : TracerPourcentage(g)
            Case IconeWU.Enregistrer : TracerEnregistrer(g)
            Case IconeWU.Fermer : TracerFermer(g)
            Case IconeWU.Nouveau : TracerNouveau(g)
            Case IconeWU.Modifier : TracerModifier(g)
            Case IconeWU.Supprimer : TracerSupprimer(g)
            Case IconeWU.Actualiser : TracerActualiser(g)
            Case IconeWU.Exporter : TracerExporter(g)
            Case IconeWU.Copier : TracerCopier(g)
            Case IconeWU.Annuler : TracerAnnuler(g)
            Case IconeWU.Calculer : TracerCalculer(g)
            Case IconeWU.Telecharger : TracerTelecharger(g)
            Case IconeWU.Calendrier : TracerCalendrier(g)
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
    ''' <summary>
    ''' Le barème des taxes : un signe pour cent, dessiné à la main plutôt qu'écrit au
    ''' clavier. Un caractère posé avec DrawString changerait de forme d'un poste à l'autre
    ''' selon les polices installées, et se placerait mal aux petites tailles ; deux disques
    ''' et une barre oblique tiennent leur place à seize pixels comme à soixante-quatre.
    ''' </summary>
    ''' <summary>Disquette : enregistrer. Le volet clair en haut, l'étiquette en bas.</summary>
    Private Shared Sub TracerEnregistrer(g As Graphics)

        Boite(g, 2.0F, 2.5F, 12.0F, 11.0F, TEINTE_BLEU)

        ' L'obturateur, en haut : c'est lui qui fait reconnaître la disquette à 16 pixels,
        ' bien plus que le carré qui la porte.
        Boite(g, 5.0F, 3.0F, 6.0F, 4.0F, TEINTE_PAPIER)
        Boite(g, 8.6F, 3.6F, 1.6F, 2.8F, TEINTE_ARDOISE, False)

        ' L'étiquette, en bas.
        Boite(g, 4.0F, 9.0F, 8.0F, 4.5F, TEINTE_PAPIER)
    End Sub

    ''' <summary>Croix : fermer la fenêtre. Ardoise, et non rouge — fermer n'est pas annuler.</summary>
    Private Shared Sub TracerFermer(g As Graphics)

        Disque(g, 8.0F, 8.0F, 6.0F, TEINTE_PAPIER)

        Using stylo As Pen = Contour(1.8F)
            g.DrawLine(stylo, 5.4F, 5.4F, 10.6F, 10.6F)
            g.DrawLine(stylo, 10.6F, 5.4F, 5.4F, 10.6F)
        End Using
    End Sub

    ''' <summary>Feuille et signe plus : créer un élément.</summary>
    Private Shared Sub TracerNouveau(g As Graphics)

        Boite(g, 3.0F, 1.5F, 8.0F, 11.0F, TEINTE_PAPIER)

        Using stylo As Pen = Contour(0.9F)
            g.DrawLine(stylo, 4.8F, 4.5F, 9.2F, 4.5F)
            g.DrawLine(stylo, 4.8F, 6.5F, 9.2F, 6.5F)
        End Using

        ' Le plus est posé par-dessus, en bas à droite : il déborde de la feuille, sinon il se
        ' confond avec une ligne de texte.
        Disque(g, 11.5F, 11.5F, 4.0F, TEINTE_VERT)

        Using stylo As New Pen(TEINTE_PAPIER, 1.6F)
            g.DrawLine(stylo, 9.4F, 11.5F, 13.6F, 11.5F)
            g.DrawLine(stylo, 11.5F, 9.4F, 11.5F, 13.6F)
        End Using
    End Sub

    ''' <summary>Crayon en diagonale : modifier.</summary>
    Private Shared Sub TracerModifier(g As Graphics)

        ' Le corps, de la pointe en bas à gauche vers la gomme en haut à droite.
        Forme(g, New PointF() {New PointF(5.0F, 13.0F), New PointF(3.0F, 13.5F),
                               New PointF(3.5F, 11.5F), New PointF(11.5F, 3.5F),
                               New PointF(13.0F, 5.0F)}, TEINTE_OR)

        ' La pointe, plus sombre : c'est elle qui donne le sens du crayon.
        Forme(g, New PointF() {New PointF(3.0F, 13.5F), New PointF(3.5F, 11.5F),
                               New PointF(5.0F, 13.0F)}, TEINTE_ARDOISE)

        ' La virole, en haut.
        Forme(g, New PointF() {New PointF(11.5F, 3.5F), New PointF(12.5F, 2.5F),
                               New PointF(14.0F, 4.0F), New PointF(13.0F, 5.0F)}, TEINTE_BLEU)
    End Sub

    ''' <summary>Corbeille : supprimer. La seule icône d'action en teinte d'alerte.</summary>
    Private Shared Sub TracerSupprimer(g As Graphics)

        ' Le couvercle et sa poignée.
        Boite(g, 2.5F, 4.0F, 11.0F, 1.8F, TEINTE_ALERTE)
        Boite(g, 6.2F, 2.2F, 3.6F, 1.8F, TEINTE_ALERTE)

        ' Le corps, légèrement tronconique.
        Forme(g, New PointF() {New PointF(4.0F, 6.2F), New PointF(12.0F, 6.2F),
                               New PointF(11.0F, 14.0F), New PointF(5.0F, 14.0F)}, TEINTE_ALERTE)

        Using stylo As New Pen(TEINTE_PAPIER, 0.9F)
            g.DrawLine(stylo, 6.6F, 8.0F, 6.4F, 12.2F)
            g.DrawLine(stylo, 8.0F, 8.0F, 8.0F, 12.2F)
            g.DrawLine(stylo, 9.4F, 8.0F, 9.6F, 12.2F)
        End Using
    End Sub

    ''' <summary>Flèche circulaire : actualiser, resynchroniser.</summary>
    Private Shared Sub TracerActualiser(g As Graphics)

        Using stylo As New Pen(TEINTE_BLEU, 2.0F)
            stylo.StartCap = LineCap.Round
            stylo.EndCap = LineCap.Round

            ' L'arc est ouvert en haut à droite : c'est cette ouverture qui fait lire une
            ' flèche plutôt qu'un cercle.
            g.DrawArc(stylo, 3.0F, 3.0F, 10.0F, 10.0F, 300.0F, 300.0F)
        End Using

        Forme(g, New PointF() {New PointF(11.0F, 1.2F), New PointF(14.2F, 4.2F),
                               New PointF(10.0F, 5.2F)}, TEINTE_BLEU, False)
    End Sub

    ''' <summary>Feuille et flèche sortante : exporter, imprimer un document.</summary>
    Private Shared Sub TracerExporter(g As Graphics)

        Boite(g, 2.0F, 1.5F, 8.0F, 11.0F, TEINTE_PAPIER)

        Using stylo As Pen = Contour(0.9F)
            g.DrawLine(stylo, 3.8F, 4.2F, 8.2F, 4.2F)
            g.DrawLine(stylo, 3.8F, 6.2F, 8.2F, 6.2F)
            g.DrawLine(stylo, 3.8F, 8.2F, 6.5F, 8.2F)
        End Using

        ' La flèche sort de la feuille vers la droite : c'est le sens qui distingue l'export
        ' de l'import, et non la couleur.
        Using stylo As New Pen(TEINTE_VERT, 1.8F)
            stylo.StartCap = LineCap.Round
            g.DrawLine(stylo, 8.0F, 11.5F, 12.5F, 11.5F)
        End Using

        Forme(g, New PointF() {New PointF(12.0F, 9.2F), New PointF(15.0F, 11.5F),
                               New PointF(12.0F, 13.8F)}, TEINTE_VERT, False)
    End Sub

    ''' <summary>Deux feuilles décalées : copier.</summary>
    Private Shared Sub TracerCopier(g As Graphics)

        Boite(g, 2.0F, 1.5F, 7.5F, 9.5F, TEINTE_PAPIER)
        Boite(g, 6.5F, 5.0F, 7.5F, 9.5F, TEINTE_BLEU_CLAIR)

        Using stylo As New Pen(TEINTE_PAPIER, 0.9F)
            g.DrawLine(stylo, 8.2F, 7.8F, 12.3F, 7.8F)
            g.DrawLine(stylo, 8.2F, 10.0F, 12.3F, 10.0F)
            g.DrawLine(stylo, 8.2F, 12.2F, 10.5F, 12.2F)
        End Using
    End Sub

    ''' <summary>Flèche de retour : annuler, renoncer, rejeter.</summary>
    Private Shared Sub TracerAnnuler(g As Graphics)

        Using stylo As New Pen(TEINTE_ARDOISE, 1.9F)
            stylo.EndCap = LineCap.Round
            g.DrawArc(stylo, 3.5F, 4.0F, 10.0F, 9.0F, 200.0F, 250.0F)
        End Using

        ' La pointe tournée vers la gauche : on revient en arrière.
        Forme(g, New PointF() {New PointF(1.5F, 6.5F), New PointF(6.5F, 5.0F),
                               New PointF(5.2F, 9.8F)}, TEINTE_ARDOISE, False)
    End Sub

    ''' <summary>Calculatrice : afficher et calculer la journée.</summary>
    Private Shared Sub TracerCalculer(g As Graphics)

        Boite(g, 2.5F, 1.5F, 11.0F, 13.0F, TEINTE_ARDOISE)

        ' L'écran, en haut.
        Boite(g, 4.0F, 3.0F, 8.0F, 3.0F, TEINTE_VERT, False)

        ' SIX TOUCHES, POSÉES UNE PAR UNE. Elles l'étaient par deux boucles imbriquées, ce qui
        ' se lisait mieux — mais l'outil qui rend ces dessins en PNG n'interprète que les
        ' appels littéraux : la calculatrice sortait en ardoise unie sur la planche de
        ' contrôle, et le seul moyen de la regarder avant de livrer était perdu.
        '
        ' Neuf touches donnaient d'ailleurs des carrés d'un pixel et demi. Six se voient.
        Boite(g, 4.0F, 8.0F, 2.1F, 2.2F, TEINTE_PAPIER, False)
        Boite(g, 6.9F, 8.0F, 2.1F, 2.2F, TEINTE_PAPIER, False)
        Boite(g, 9.8F, 8.0F, 2.1F, 2.2F, TEINTE_PAPIER, False)
        Boite(g, 4.0F, 11.2F, 2.1F, 2.2F, TEINTE_PAPIER, False)
        Boite(g, 6.9F, 11.2F, 2.1F, 2.2F, TEINTE_PAPIER, False)
        Boite(g, 9.8F, 11.2F, 2.1F, 2.2F, TEINTE_PAPIER, False)
    End Sub

    ''' <summary>Flèche descendante sur un bac : charger un rapport depuis un fichier.</summary>
    Private Shared Sub TracerTelecharger(g As Graphics)

        Using stylo As New Pen(TEINTE_BLEU, 2.0F)
            stylo.StartCap = LineCap.Round
            g.DrawLine(stylo, 8.0F, 1.5F, 8.0F, 7.5F)
        End Using

        Forme(g, New PointF() {New PointF(4.8F, 6.8F), New PointF(11.2F, 6.8F),
                               New PointF(8.0F, 10.5F)}, TEINTE_BLEU, False)

        ' Le bac : deux jambages et un fond. Sans lui, la flèche seule se lit « descendre ».
        Using stylo As Pen = Contour(1.6F)
            g.DrawLine(stylo, 2.5F, 10.5F, 2.5F, 13.5F)
            g.DrawLine(stylo, 2.5F, 13.5F, 13.5F, 13.5F)
            g.DrawLine(stylo, 13.5F, 13.5F, 13.5F, 10.5F)
        End Using
    End Sub

    ''' <summary>Calendrier : la période complète d'un mois.</summary>
    Private Shared Sub TracerCalendrier(g As Graphics)

        Boite(g, 2.0F, 3.0F, 12.0F, 11.0F, TEINTE_PAPIER)
        Boite(g, 2.0F, 3.0F, 12.0F, 3.0F, TEINTE_OR)

        ' Les deux anneaux, qui débordent du bandeau.
        Using stylo As Pen = Contour(1.4F)
            g.DrawLine(stylo, 5.0F, 1.5F, 5.0F, 4.0F)
            g.DrawLine(stylo, 11.0F, 1.5F, 11.0F, 4.0F)
        End Using

        ' Trois jours marqués : assez pour lire un calendrier, trop peu pour faire une grille
        ' illisible à seize pixels.
        Boite(g, 4.0F, 8.0F, 2.0F, 2.0F, TEINTE_BLEU_CLAIR, False)
        Boite(g, 7.0F, 8.0F, 2.0F, 2.0F, TEINTE_BLEU_CLAIR, False)
        Boite(g, 10.0F, 8.0F, 2.0F, 2.0F, TEINTE_BLEU_CLAIR, False)
        Boite(g, 4.0F, 11.0F, 2.0F, 2.0F, TEINTE_BLEU_CLAIR, False)
    End Sub

    Private Shared Sub TracerPourcentage(g As Graphics)

        Boite(g, 2.2F, 2.2F, 11.6F, 11.6F, TEINTE_PAPIER)

        ' La barre d'abord : les deux disques viennent ensuite s'y poser, et la recouvrent
        ' proprement là où elle les traverserait.
        Using barre As New Pen(TEINTE_ARDOISE, 1.6F)
            barre.StartCap = LineCap.Round
            barre.EndCap = LineCap.Round
            g.DrawLine(barre, 4.6F, 11.4F, 11.4F, 4.6F)
        End Using

        Disque(g, 5.4F, 5.4F, 1.9F, TEINTE_OR)
        Disque(g, 10.6F, 10.6F, 1.9F, TEINTE_BLEU)
    End Sub

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

    ''' <summary>Aucune icône : un bouton dont le nom n'est pas connu reste nu.</summary>
    Aucune = 0

    ' Les vingt dessins d'ÉCRAN : ils nomment une fenêtre, et habillent les menus.
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
    Pourcentage = 20

    ' Les douze dessins d'ACTION, posés sur les boutons. Les vingt précédents nomment des
    ' écrans ; ceux-ci nomment des VERBES — enregistrer, fermer, supprimer. La distinction
    ' n'est pas décorative : un bouton porte ce qu'il FAIT, un menu ce qu'il OUVRE.
    Enregistrer = 21
    Fermer = 22
    Nouveau = 23
    Modifier = 24
    Supprimer = 25
    Actualiser = 26
    Exporter = 27
    Copier = 28
    Annuler = 29
    Calculer = 30
    Telecharger = 31
    Calendrier = 32
End Enum
