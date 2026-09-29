Option Strict On
Option Explicit On

Partial Class FrmProduits
    Inherits System.Windows.Forms.Form

    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    Private components As System.ComponentModel.IContainer

    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.lblTitre = New System.Windows.Forms.Label()
        Me.lblSousTitre = New System.Windows.Forms.Label()
        Me.lvProduits = New System.Windows.Forms.ListView()
        Me.colCode = New System.Windows.Forms.ColumnHeader()
        Me.colNom = New System.Windows.Forms.ColumnHeader()
        Me.colOrdre = New System.Windows.Forms.ColumnHeader()
        Me.colEtat = New System.Windows.Forms.ColumnHeader()
        Me.imgLogos = New System.Windows.Forms.ImageList(Me.components)
        Me.grpProduit = New System.Windows.Forms.GroupBox()
        Me.lblCode = New System.Windows.Forms.Label()
        Me.txtCode = New System.Windows.Forms.TextBox()
        Me.lblRegleDuCode = New System.Windows.Forms.Label()
        Me.lblNom = New System.Windows.Forms.Label()
        Me.txtNom = New System.Windows.Forms.TextBox()
        Me.lblDescription = New System.Windows.Forms.Label()
        Me.txtDescription = New System.Windows.Forms.TextBox()
        Me.lblCouleur = New System.Windows.Forms.Label()
        Me.pnlCouleur = New System.Windows.Forms.Panel()
        Me.btnCouleur = New System.Windows.Forms.Button()
        Me.lblOrdre = New System.Windows.Forms.Label()
        Me.numOrdre = New System.Windows.Forms.NumericUpDown()
        Me.chkEnService = New System.Windows.Forms.CheckBox()
        Me.lblEtat = New System.Windows.Forms.Label()
        Me.lblTracabilite = New System.Windows.Forms.Label()
        Me.btnNouveau = New System.Windows.Forms.Button()
        Me.btnEnregistrer = New System.Windows.Forms.Button()
        Me.btnFermer = New System.Windows.Forms.Button()
        Me.lblStatut = New System.Windows.Forms.Label()
        Me.dlgCouleur = New System.Windows.Forms.ColorDialog()
        Me.grpProduit.SuspendLayout()
        CType(Me.numOrdre, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'lblTitre
        '
        Me.lblTitre.AutoSize = True
        Me.lblTitre.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Bold)
        Me.lblTitre.Location = New System.Drawing.Point(16, 14)
        Me.lblTitre.Name = "lblTitre"
        Me.lblTitre.Size = New System.Drawing.Size(240, 21)
        Me.lblTitre.TabIndex = 0
        Me.lblTitre.Text = "Produits de transfert"
        '
        'lblSousTitre
        '
        Me.lblSousTitre.ForeColor = System.Drawing.SystemColors.GrayText
        Me.lblSousTitre.Location = New System.Drawing.Point(18, 40)
        Me.lblSousTitre.Name = "lblSousTitre"
        Me.lblSousTitre.Size = New System.Drawing.Size(840, 36)
        Me.lblSousTitre.TabIndex = 1
        Me.lblSousTitre.Text = "Cette liste dit quels produits EXISTENT. C'est l'application qui sait lesquels ell" & _
            "e sait TRAITER : un produit ajouté ici reste « En attente » tant qu'une livraison" & _
            " n'apporte pas le traitement de sa compensation."
        '
        'lvProduits
        '
        Me.lvProduits.FullRowSelect = True
        Me.lvProduits.GridLines = True
        Me.lvProduits.HideSelection = False
        Me.lvProduits.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.colCode, Me.colNom, Me.colOrdre, Me.colEtat})
        Me.lvProduits.Location = New System.Drawing.Point(18, 84)
        Me.lvProduits.MultiSelect = False
        Me.lvProduits.Name = "lvProduits"
        Me.lvProduits.Size = New System.Drawing.Size(840, 180)
        Me.lvProduits.SmallImageList = Me.imgLogos
        Me.lvProduits.TabIndex = 2
        Me.lvProduits.UseCompatibleStateImageBehavior = False
        Me.lvProduits.View = System.Windows.Forms.View.Details
        '
        'colCode
        '
        Me.colCode.Text = "Code"
        Me.colCode.Width = 140
        '
        'colNom
        '
        Me.colNom.Text = "Nom"
        Me.colNom.Width = 360
        '
        'colOrdre
        '
        Me.colOrdre.Text = "Ordre"
        Me.colOrdre.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.colOrdre.Width = 80
        '
        'colEtat
        '
        Me.colEtat.Text = "État"
        Me.colEtat.Width = 220
        '
        'imgLogos
        '
        Me.imgLogos.ColorDepth = System.Windows.Forms.ColorDepth.Depth32Bit
        Me.imgLogos.ImageSize = New System.Drawing.Size(24, 24)
        Me.imgLogos.TransparentColor = System.Drawing.Color.Transparent
        '
        'grpProduit
        '
        Me.grpProduit.Controls.Add(Me.lblTracabilite)
        Me.grpProduit.Controls.Add(Me.lblEtat)
        Me.grpProduit.Controls.Add(Me.chkEnService)
        Me.grpProduit.Controls.Add(Me.numOrdre)
        Me.grpProduit.Controls.Add(Me.lblOrdre)
        Me.grpProduit.Controls.Add(Me.btnCouleur)
        Me.grpProduit.Controls.Add(Me.pnlCouleur)
        Me.grpProduit.Controls.Add(Me.lblCouleur)
        Me.grpProduit.Controls.Add(Me.txtDescription)
        Me.grpProduit.Controls.Add(Me.lblDescription)
        Me.grpProduit.Controls.Add(Me.txtNom)
        Me.grpProduit.Controls.Add(Me.lblNom)
        Me.grpProduit.Controls.Add(Me.lblRegleDuCode)
        Me.grpProduit.Controls.Add(Me.txtCode)
        Me.grpProduit.Controls.Add(Me.lblCode)
        Me.grpProduit.Location = New System.Drawing.Point(18, 278)
        Me.grpProduit.Name = "grpProduit"
        Me.grpProduit.Size = New System.Drawing.Size(840, 262)
        Me.grpProduit.TabIndex = 3
        Me.grpProduit.TabStop = False
        Me.grpProduit.Text = "Le produit"
        '
        'lblCode
        '
        Me.lblCode.AutoSize = True
        Me.lblCode.Location = New System.Drawing.Point(18, 30)
        Me.lblCode.Name = "lblCode"
        Me.lblCode.Size = New System.Drawing.Size(40, 15)
        Me.lblCode.TabIndex = 0
        Me.lblCode.Text = "&Code"
        '
        'txtCode
        '
        Me.txtCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtCode.Location = New System.Drawing.Point(150, 26)
        Me.txtCode.MaxLength = 10
        Me.txtCode.Name = "txtCode"
        Me.txtCode.Size = New System.Drawing.Size(120, 23)
        Me.txtCode.TabIndex = 1
        '
        'lblRegleDuCode
        '
        Me.lblRegleDuCode.ForeColor = System.Drawing.SystemColors.GrayText
        Me.lblRegleDuCode.Location = New System.Drawing.Point(284, 22)
        Me.lblRegleDuCode.Name = "lblRegleDuCode"
        Me.lblRegleDuCode.Size = New System.Drawing.Size(538, 34)
        Me.lblRegleDuCode.TabIndex = 2
        Me.lblRegleDuCode.Text = "2 à 10 lettres non accentuées et chiffres, commençant par une lettre. Il nomme le " & _
            "fichier du logo (Logos\CODE.png) et, demain, les tables du produit : il ne se mod" & _
            "ifie plus après création."
        '
        'lblNom
        '
        Me.lblNom.AutoSize = True
        Me.lblNom.Location = New System.Drawing.Point(18, 70)
        Me.lblNom.Name = "lblNom"
        Me.lblNom.Size = New System.Drawing.Size(35, 15)
        Me.lblNom.TabIndex = 3
        Me.lblNom.Text = "&Nom"
        '
        'txtNom
        '
        Me.txtNom.Location = New System.Drawing.Point(150, 66)
        Me.txtNom.MaxLength = 100
        Me.txtNom.Name = "txtNom"
        Me.txtNom.Size = New System.Drawing.Size(360, 23)
        Me.txtNom.TabIndex = 4
        '
        'lblDescription
        '
        Me.lblDescription.AutoSize = True
        Me.lblDescription.Location = New System.Drawing.Point(18, 104)
        Me.lblDescription.Name = "lblDescription"
        Me.lblDescription.Size = New System.Drawing.Size(70, 15)
        Me.lblDescription.TabIndex = 5
        Me.lblDescription.Text = "&Description"
        '
        'txtDescription
        '
        Me.txtDescription.Location = New System.Drawing.Point(150, 100)
        Me.txtDescription.MaxLength = 500
        Me.txtDescription.Multiline = True
        Me.txtDescription.Name = "txtDescription"
        Me.txtDescription.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtDescription.Size = New System.Drawing.Size(672, 60)
        Me.txtDescription.TabIndex = 6
        '
        'lblCouleur
        '
        Me.lblCouleur.AutoSize = True
        Me.lblCouleur.Location = New System.Drawing.Point(18, 178)
        Me.lblCouleur.Name = "lblCouleur"
        Me.lblCouleur.Size = New System.Drawing.Size(52, 15)
        Me.lblCouleur.TabIndex = 7
        Me.lblCouleur.Text = "Cou&leur"
        '
        'pnlCouleur
        '
        Me.pnlCouleur.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlCouleur.Location = New System.Drawing.Point(150, 174)
        Me.pnlCouleur.Name = "pnlCouleur"
        Me.pnlCouleur.Size = New System.Drawing.Size(60, 24)
        Me.pnlCouleur.TabIndex = 8
        '
        'btnCouleur
        '
        Me.btnCouleur.Location = New System.Drawing.Point(222, 172)
        Me.btnCouleur.Name = "btnCouleur"
        Me.btnCouleur.Size = New System.Drawing.Size(120, 28)
        Me.btnCouleur.TabIndex = 9
        Me.btnCouleur.Text = "C&hoisir..."
        Me.btnCouleur.UseVisualStyleBackColor = True
        '
        'lblOrdre
        '
        Me.lblOrdre.AutoSize = True
        Me.lblOrdre.Location = New System.Drawing.Point(372, 178)
        Me.lblOrdre.Name = "lblOrdre"
        Me.lblOrdre.Size = New System.Drawing.Size(42, 15)
        Me.lblOrdre.TabIndex = 10
        Me.lblOrdre.Text = "&Ordre"
        '
        'numOrdre
        '
        Me.numOrdre.Location = New System.Drawing.Point(430, 174)
        Me.numOrdre.Maximum = New Decimal(New Integer() {9999, 0, 0, 0})
        Me.numOrdre.Name = "numOrdre"
        Me.numOrdre.Size = New System.Drawing.Size(80, 23)
        Me.numOrdre.TabIndex = 11
        '
        'chkEnService
        '
        Me.chkEnService.AutoSize = True
        Me.chkEnService.Location = New System.Drawing.Point(566, 176)
        Me.chkEnService.Name = "chkEnService"
        Me.chkEnService.Size = New System.Drawing.Size(90, 19)
        Me.chkEnService.TabIndex = 12
        Me.chkEnService.Text = "En &service"
        Me.chkEnService.UseVisualStyleBackColor = True
        '
        'lblEtat
        '
        Me.lblEtat.Location = New System.Drawing.Point(147, 210)
        Me.lblEtat.Name = "lblEtat"
        Me.lblEtat.Size = New System.Drawing.Size(675, 20)
        Me.lblEtat.TabIndex = 13
        Me.lblEtat.Text = ""
        '
        'lblTracabilite
        '
        Me.lblTracabilite.ForeColor = System.Drawing.SystemColors.GrayText
        Me.lblTracabilite.Location = New System.Drawing.Point(147, 232)
        Me.lblTracabilite.Name = "lblTracabilite"
        Me.lblTracabilite.Size = New System.Drawing.Size(675, 20)
        Me.lblTracabilite.TabIndex = 14
        Me.lblTracabilite.Text = ""
        '
        'btnNouveau
        '
        Me.btnNouveau.Location = New System.Drawing.Point(18, 556)
        Me.btnNouveau.Name = "btnNouveau"
        Me.btnNouveau.Size = New System.Drawing.Size(150, 32)
        Me.btnNouveau.TabIndex = 4
        Me.btnNouveau.Text = "Nouveau &produit"
        Me.btnNouveau.UseVisualStyleBackColor = True
        '
        'btnEnregistrer
        '
        Me.btnEnregistrer.Location = New System.Drawing.Point(600, 556)
        Me.btnEnregistrer.Name = "btnEnregistrer"
        Me.btnEnregistrer.Size = New System.Drawing.Size(140, 32)
        Me.btnEnregistrer.TabIndex = 5
        Me.btnEnregistrer.Text = "&Enregistrer"
        Me.btnEnregistrer.UseVisualStyleBackColor = True
        '
        'btnFermer
        '
        Me.btnFermer.Location = New System.Drawing.Point(750, 556)
        Me.btnFermer.Name = "btnFermer"
        Me.btnFermer.Size = New System.Drawing.Size(108, 32)
        Me.btnFermer.TabIndex = 6
        Me.btnFermer.Text = "&Fermer"
        Me.btnFermer.UseVisualStyleBackColor = True
        '
        'lblStatut
        '
        Me.lblStatut.ForeColor = System.Drawing.SystemColors.GrayText
        Me.lblStatut.Location = New System.Drawing.Point(18, 598)
        Me.lblStatut.Name = "lblStatut"
        Me.lblStatut.Size = New System.Drawing.Size(840, 36)
        Me.lblStatut.TabIndex = 7
        Me.lblStatut.Text = ""
        '
        'FrmProduits
        '
        Me.AcceptButton = Me.btnEnregistrer
        Me.CancelButton = Me.btnFermer
        Me.ClientSize = New System.Drawing.Size(876, 644)
        Me.Controls.Add(Me.lblStatut)
        Me.Controls.Add(Me.btnFermer)
        Me.Controls.Add(Me.btnEnregistrer)
        Me.Controls.Add(Me.btnNouveau)
        Me.Controls.Add(Me.grpProduit)
        Me.Controls.Add(Me.lvProduits)
        Me.Controls.Add(Me.lblSousTitre)
        Me.Controls.Add(Me.lblTitre)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.Name = "FrmProduits"
        Me.Text = "Produits de transfert"
        Me.grpProduit.ResumeLayout(False)
        Me.grpProduit.PerformLayout()
        CType(Me.numOrdre, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents lblTitre As System.Windows.Forms.Label
    Friend WithEvents lblSousTitre As System.Windows.Forms.Label
    Friend WithEvents lvProduits As System.Windows.Forms.ListView
    Friend WithEvents colCode As System.Windows.Forms.ColumnHeader
    Friend WithEvents colNom As System.Windows.Forms.ColumnHeader
    Friend WithEvents colOrdre As System.Windows.Forms.ColumnHeader
    Friend WithEvents colEtat As System.Windows.Forms.ColumnHeader
    Friend WithEvents imgLogos As System.Windows.Forms.ImageList
    Friend WithEvents grpProduit As System.Windows.Forms.GroupBox
    Friend WithEvents lblCode As System.Windows.Forms.Label
    Friend WithEvents txtCode As System.Windows.Forms.TextBox
    Friend WithEvents lblRegleDuCode As System.Windows.Forms.Label
    Friend WithEvents lblNom As System.Windows.Forms.Label
    Friend WithEvents txtNom As System.Windows.Forms.TextBox
    Friend WithEvents lblDescription As System.Windows.Forms.Label
    Friend WithEvents txtDescription As System.Windows.Forms.TextBox
    Friend WithEvents lblCouleur As System.Windows.Forms.Label
    Friend WithEvents pnlCouleur As System.Windows.Forms.Panel
    Friend WithEvents btnCouleur As System.Windows.Forms.Button
    Friend WithEvents lblOrdre As System.Windows.Forms.Label
    Friend WithEvents numOrdre As System.Windows.Forms.NumericUpDown
    Friend WithEvents chkEnService As System.Windows.Forms.CheckBox
    Friend WithEvents lblEtat As System.Windows.Forms.Label
    Friend WithEvents lblTracabilite As System.Windows.Forms.Label
    Friend WithEvents btnNouveau As System.Windows.Forms.Button
    Friend WithEvents btnEnregistrer As System.Windows.Forms.Button
    Friend WithEvents btnFermer As System.Windows.Forms.Button
    Friend WithEvents lblStatut As System.Windows.Forms.Label
    Friend WithEvents dlgCouleur As System.Windows.Forms.ColorDialog

End Class
