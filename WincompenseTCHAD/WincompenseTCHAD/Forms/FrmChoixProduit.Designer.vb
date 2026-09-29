Option Strict On
Option Explicit On

Partial Class FrmChoixProduit
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
        Me.imgLogos = New System.Windows.Forms.ImageList(Me.components)
        Me.picLogo = New System.Windows.Forms.PictureBox()
        Me.tipLogo = New System.Windows.Forms.ToolTip(Me.components)
        Me.colProduit = New System.Windows.Forms.ColumnHeader()
        Me.colEtat = New System.Windows.Forms.ColumnHeader()
        Me.lblDescription = New System.Windows.Forms.Label()
        Me.lblUtilisateur = New System.Windows.Forms.Label()
        Me.btnOuvrir = New System.Windows.Forms.Button()
        Me.btnQuitter = New System.Windows.Forms.Button()
        CType(Me.picLogo, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'lblTitre
        '
        Me.lblTitre.AutoSize = True
        Me.lblTitre.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Bold)
        Me.lblTitre.Location = New System.Drawing.Point(16, 14)
        Me.lblTitre.Name = "lblTitre"
        Me.lblTitre.Size = New System.Drawing.Size(220, 21)
        Me.lblTitre.TabIndex = 0
        Me.lblTitre.Text = "Produit de transfert"
        '
        'lblSousTitre
        '
        Me.lblSousTitre.AutoSize = True
        Me.lblSousTitre.ForeColor = System.Drawing.SystemColors.GrayText
        Me.lblSousTitre.Location = New System.Drawing.Point(18, 42)
        Me.lblSousTitre.Name = "lblSousTitre"
        Me.lblSousTitre.Size = New System.Drawing.Size(300, 15)
        Me.lblSousTitre.TabIndex = 1
        Me.lblSousTitre.Text = "Choisissez le produit dont vous voulez traiter la compensation."
        '
        'lvProduits
        '
        Me.lvProduits.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lvProduits.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.colProduit, Me.colEtat})
        Me.lvProduits.FullRowSelect = True
        Me.lvProduits.GridLines = True
        Me.lvProduits.HideSelection = False
        Me.lvProduits.Location = New System.Drawing.Point(18, 70)
        Me.lvProduits.MultiSelect = False
        Me.lvProduits.Name = "lvProduits"
        Me.lvProduits.Size = New System.Drawing.Size(494, 170)
        Me.lvProduits.SmallImageList = Me.imgLogos
        Me.lvProduits.TabIndex = 2
        Me.lvProduits.UseCompatibleStateImageBehavior = False
        Me.lvProduits.View = System.Windows.Forms.View.Details
        '
        'imgLogos
        '
        ' Trente-deux pixels : la hauteur d'une ligne de liste avec une marge, et une taille ou
        ' un logo reste identifiable. ColorDepth en 32 bits pour la transparence des PNG.
        Me.imgLogos.ColorDepth = System.Windows.Forms.ColorDepth.Depth32Bit
        Me.imgLogos.ImageSize = New System.Drawing.Size(32, 32)
        Me.imgLogos.TransparentColor = System.Drawing.Color.Transparent
        '
        'colProduit
        '
        Me.colProduit.Text = "Produit de transfert"
        Me.colProduit.Width = 320
        '
        'colEtat
        '
        Me.colEtat.Text = "État"
        Me.colEtat.Width = 150
        '
        'lblDescription
        '
        Me.lblDescription.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblDescription.Location = New System.Drawing.Point(130, 256)
        Me.lblDescription.Name = "lblDescription"
        Me.lblDescription.Size = New System.Drawing.Size(382, 96)
        Me.lblDescription.TabIndex = 4
        Me.lblDescription.Text = ""
        '
        'picLogo
        '
        ' CenterImage et non Zoom : l'image est deja rendue exactement a la taille du cadre,
        ' proportions conservees. Laisser le cadre la remettre a l'echelle la ferait passer
        ' deux fois par un reechantillonnage, et un logo fin y perdrait ses traits.
        Me.picLogo.Location = New System.Drawing.Point(18, 256)
        Me.picLogo.Name = "picLogo"
        Me.picLogo.Size = New System.Drawing.Size(96, 96)
        Me.picLogo.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.picLogo.TabIndex = 3
        Me.picLogo.TabStop = False
        '
        'lblUtilisateur
        '
        Me.lblUtilisateur.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        ' Pas d'AutoSize : la description d'une session -- nom complet, identifiant, role et
        ' fonction -- est longue, et une etiquette qui grandit passerait sous les boutons.
        Me.lblUtilisateur.AutoEllipsis = True
        Me.lblUtilisateur.ForeColor = System.Drawing.SystemColors.GrayText
        Me.lblUtilisateur.Location = New System.Drawing.Point(18, 368)
        Me.lblUtilisateur.Name = "lblUtilisateur"
        Me.lblUtilisateur.Size = New System.Drawing.Size(270, 30)
        Me.lblUtilisateur.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.lblUtilisateur.TabIndex = 5
        Me.lblUtilisateur.Text = ""
        '
        'btnOuvrir
        '
        Me.btnOuvrir.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOuvrir.Location = New System.Drawing.Point(300, 368)
        Me.btnOuvrir.Name = "btnOuvrir"
        Me.btnOuvrir.Size = New System.Drawing.Size(120, 30)
        Me.btnOuvrir.TabIndex = 6
        Me.btnOuvrir.Text = "&Ouvrir"
        Me.btnOuvrir.UseVisualStyleBackColor = True
        '
        'btnQuitter
        '
        Me.btnQuitter.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnQuitter.Location = New System.Drawing.Point(430, 368)
        Me.btnQuitter.Name = "btnQuitter"
        Me.btnQuitter.Size = New System.Drawing.Size(82, 30)
        Me.btnQuitter.TabIndex = 7
        Me.btnQuitter.Text = "&Quitter"
        Me.btnQuitter.UseVisualStyleBackColor = True
        '
        'FrmChoixProduit
        '
        Me.AcceptButton = Me.btnOuvrir
        Me.CancelButton = Me.btnQuitter
        Me.ClientSize = New System.Drawing.Size(530, 414)
        Me.Controls.Add(Me.btnQuitter)
        Me.Controls.Add(Me.btnOuvrir)
        Me.Controls.Add(Me.lblUtilisateur)
        Me.Controls.Add(Me.picLogo)
        Me.Controls.Add(Me.lblDescription)
        Me.Controls.Add(Me.lvProduits)
        Me.Controls.Add(Me.lblSousTitre)
        Me.Controls.Add(Me.lblTitre)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FrmChoixProduit"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Wincompense TCHAD"
        CType(Me.picLogo, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents lblTitre As System.Windows.Forms.Label
    Friend WithEvents lblSousTitre As System.Windows.Forms.Label
    Friend WithEvents lvProduits As System.Windows.Forms.ListView
    Friend WithEvents colProduit As System.Windows.Forms.ColumnHeader
    Friend WithEvents colEtat As System.Windows.Forms.ColumnHeader
    Friend WithEvents imgLogos As System.Windows.Forms.ImageList
    Friend WithEvents picLogo As System.Windows.Forms.PictureBox
    Friend WithEvents tipLogo As System.Windows.Forms.ToolTip
    Friend WithEvents lblDescription As System.Windows.Forms.Label
    Friend WithEvents lblUtilisateur As System.Windows.Forms.Label
    Friend WithEvents btnOuvrir As System.Windows.Forms.Button
    Friend WithEvents btnQuitter As System.Windows.Forms.Button

End Class
