Option Strict On
Option Explicit On

Partial Class FrmControleChange
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
        Me.panelTitre = New System.Windows.Forms.Panel()
        Me.lblTitre = New System.Windows.Forms.Label()
        Me.lblIntro = New System.Windows.Forms.Label()
        Me.grpSource = New System.Windows.Forms.GroupBox()
        Me.btnCharger = New System.Windows.Forms.Button()
        Me.lblFichier = New System.Windows.Forms.Label()
        Me.lblParite = New System.Windows.Forms.Label()
        Me.txtParite = New System.Windows.Forms.TextBox()
        Me.chkEnAttente = New System.Windows.Forms.CheckBox()
        Me.btnAfficher = New System.Windows.Forms.Button()
        Me.onglets = New System.Windows.Forms.TabControl()
        Me.pageDetail = New System.Windows.Forms.TabPage()
        Me.dgvEcarts = New System.Windows.Forms.DataGridView()
        Me.panelFiltre = New System.Windows.Forms.Panel()
        Me.lblMtcn = New System.Windows.Forms.Label()
        Me.txtMtcn = New System.Windows.Forms.TextBox()
        Me.lblFiltre = New System.Windows.Forms.Label()
        Me.pageSynthese = New System.Windows.Forms.TabPage()
        Me.txtSynthese = New System.Windows.Forms.TextBox()
        Me.lblStatut = New System.Windows.Forms.Label()
        Me.btnCopier = New System.Windows.Forms.Button()
        Me.btnFermer = New System.Windows.Forms.Button()
        Me.ofdRapport = New System.Windows.Forms.OpenFileDialog()
        Me.panelTitre.SuspendLayout()
        Me.grpSource.SuspendLayout()
        Me.onglets.SuspendLayout()
        Me.pageDetail.SuspendLayout()
        CType(Me.dgvEcarts, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.panelFiltre.SuspendLayout()
        Me.pageSynthese.SuspendLayout()
        Me.SuspendLayout()
        '
        'panelTitre
        '
        Me.panelTitre.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.panelTitre.BackColor = System.Drawing.Color.Black
        Me.panelTitre.Controls.Add(Me.lblTitre)
        Me.panelTitre.Location = New System.Drawing.Point(12, 12)
        Me.panelTitre.Name = "panelTitre"
        Me.panelTitre.Size = New System.Drawing.Size(976, 48)
        Me.panelTitre.TabIndex = 0
        '
        'lblTitre
        '
        Me.lblTitre.BackColor = System.Drawing.Color.Black
        Me.lblTitre.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblTitre.Font = New System.Drawing.Font("Segoe UI", 14.0!, System.Drawing.FontStyle.Bold)
        Me.lblTitre.ForeColor = System.Drawing.Color.Yellow
        Me.lblTitre.Location = New System.Drawing.Point(0, 0)
        Me.lblTitre.Name = "lblTitre"
        Me.lblTitre.Size = New System.Drawing.Size(976, 48)
        Me.lblTitre.TabIndex = 0
        Me.lblTitre.Text = "Contrôle des écarts de change"
        Me.lblTitre.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblIntro
        '
        Me.lblIntro.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblIntro.Location = New System.Drawing.Point(12, 66)
        Me.lblIntro.Name = "lblIntro"
        Me.lblIntro.Size = New System.Drawing.Size(976, 36)
        Me.lblIntro.TabIndex = 1
        Me.lblIntro.Text = "Rejoue le calcul des gains et pertes de change sur un rapport de règlement, sans rien enregistrer."
        '
        'grpSource
        '
        Me.grpSource.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grpSource.Controls.Add(Me.btnCharger)
        Me.grpSource.Controls.Add(Me.lblFichier)
        Me.grpSource.Controls.Add(Me.lblParite)
        Me.grpSource.Controls.Add(Me.txtParite)
        Me.grpSource.Controls.Add(Me.chkEnAttente)
        Me.grpSource.Controls.Add(Me.btnAfficher)
        Me.grpSource.Location = New System.Drawing.Point(12, 108)
        Me.grpSource.Name = "grpSource"
        Me.grpSource.Size = New System.Drawing.Size(976, 104)
        Me.grpSource.TabIndex = 2
        Me.grpSource.TabStop = False
        Me.grpSource.Text = "Rapport de règlement et réglages du calcul"
        '
        'btnCharger
        '
        Me.btnCharger.Location = New System.Drawing.Point(12, 24)
        Me.btnCharger.Name = "btnCharger"
        Me.btnCharger.Size = New System.Drawing.Size(150, 30)
        Me.btnCharger.TabIndex = 0
        Me.btnCharger.Text = "Charger..."
        Me.btnCharger.UseVisualStyleBackColor = True
        '
        'lblFichier
        '
        Me.lblFichier.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblFichier.AutoEllipsis = True
        Me.lblFichier.ForeColor = System.Drawing.SystemColors.GrayText
        Me.lblFichier.Location = New System.Drawing.Point(174, 32)
        Me.lblFichier.Name = "lblFichier"
        Me.lblFichier.Size = New System.Drawing.Size(788, 18)
        Me.lblFichier.TabIndex = 1
        Me.lblFichier.Text = "Aucun rapport chargé."
        '
        'lblParite
        '
        Me.lblParite.Location = New System.Drawing.Point(12, 68)
        Me.lblParite.Name = "lblParite"
        Me.lblParite.Size = New System.Drawing.Size(48, 18)
        Me.lblParite.TabIndex = 2
        Me.lblParite.Text = "Parité"
        '
        'txtParite
        '
        Me.txtParite.Location = New System.Drawing.Point(66, 65)
        Me.txtParite.Name = "txtParite"
        Me.txtParite.Size = New System.Drawing.Size(96, 22)
        Me.txtParite.TabIndex = 3
        Me.txtParite.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'chkEnAttente
        '
        Me.chkEnAttente.Location = New System.Drawing.Point(182, 66)
        Me.chkEnAttente.Name = "chkEnAttente"
        Me.chkEnAttente.Size = New System.Drawing.Size(460, 22)
        Me.chkEnAttente.TabIndex = 4
        Me.chkEnAttente.Text = "Inclure les envois en attente de règlement (statut W)"
        Me.chkEnAttente.UseVisualStyleBackColor = True
        '
        'btnAfficher
        '
        Me.btnAfficher.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnAfficher.Location = New System.Drawing.Point(812, 62)
        Me.btnAfficher.Name = "btnAfficher"
        Me.btnAfficher.Size = New System.Drawing.Size(150, 30)
        Me.btnAfficher.TabIndex = 5
        Me.btnAfficher.Text = "Calculer"
        Me.btnAfficher.UseVisualStyleBackColor = True
        '
        'onglets
        '
        Me.onglets.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.onglets.Controls.Add(Me.pageDetail)
        Me.onglets.Controls.Add(Me.pageSynthese)
        Me.onglets.Location = New System.Drawing.Point(12, 222)
        Me.onglets.Name = "onglets"
        Me.onglets.SelectedIndex = 0
        Me.onglets.Size = New System.Drawing.Size(976, 350)
        Me.onglets.TabIndex = 3
        '
        'pageDetail
        '
        Me.pageDetail.Controls.Add(Me.dgvEcarts)
        Me.pageDetail.Controls.Add(Me.panelFiltre)
        Me.pageDetail.Location = New System.Drawing.Point(4, 22)
        Me.pageDetail.Name = "pageDetail"
        Me.pageDetail.Padding = New System.Windows.Forms.Padding(6)
        Me.pageDetail.Size = New System.Drawing.Size(968, 324)
        Me.pageDetail.TabIndex = 0
        Me.pageDetail.Text = "Détail des transactions"
        Me.pageDetail.UseVisualStyleBackColor = True
        '
        'dgvEcarts
        '
        Me.dgvEcarts.AllowUserToAddRows = False
        Me.dgvEcarts.AllowUserToDeleteRows = False
        Me.dgvEcarts.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvEcarts.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvEcarts.Location = New System.Drawing.Point(6, 6)
        Me.dgvEcarts.MultiSelect = False
        Me.dgvEcarts.Name = "dgvEcarts"
        Me.dgvEcarts.ReadOnly = True
        Me.dgvEcarts.RowHeadersWidth = 25
        Me.dgvEcarts.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvEcarts.Size = New System.Drawing.Size(956, 274)
        Me.dgvEcarts.TabIndex = 1
        '
        'panelFiltre
        '
        Me.panelFiltre.Controls.Add(Me.lblMtcn)
        Me.panelFiltre.Controls.Add(Me.txtMtcn)
        Me.panelFiltre.Controls.Add(Me.lblFiltre)
        Me.panelFiltre.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.panelFiltre.Location = New System.Drawing.Point(6, 280)
        Me.panelFiltre.Name = "panelFiltre"
        Me.panelFiltre.Size = New System.Drawing.Size(956, 38)
        Me.panelFiltre.TabIndex = 0
        '
        'lblMtcn
        '
        Me.lblMtcn.Location = New System.Drawing.Point(0, 10)
        Me.lblMtcn.Name = "lblMtcn"
        Me.lblMtcn.Size = New System.Drawing.Size(46, 18)
        Me.lblMtcn.TabIndex = 0
        Me.lblMtcn.Text = "MTCN"
        '
        'txtMtcn
        '
        Me.txtMtcn.Location = New System.Drawing.Point(52, 7)
        Me.txtMtcn.Name = "txtMtcn"
        Me.txtMtcn.Size = New System.Drawing.Size(140, 22)
        Me.txtMtcn.TabIndex = 1
        '
        'lblFiltre
        '
        Me.lblFiltre.ForeColor = System.Drawing.SystemColors.GrayText
        Me.lblFiltre.Location = New System.Drawing.Point(200, 10)
        Me.lblFiltre.Name = "lblFiltre"
        Me.lblFiltre.Size = New System.Drawing.Size(750, 18)
        Me.lblFiltre.TabIndex = 2
        Me.lblFiltre.Text = "Saisissez un MTCN pour ne garder que lui ; effacez pour revoir toutes les transactions."
        '
        'pageSynthese
        '
        Me.pageSynthese.Controls.Add(Me.txtSynthese)
        Me.pageSynthese.Location = New System.Drawing.Point(4, 22)
        Me.pageSynthese.Name = "pageSynthese"
        Me.pageSynthese.Padding = New System.Windows.Forms.Padding(6)
        Me.pageSynthese.Size = New System.Drawing.Size(968, 324)
        Me.pageSynthese.TabIndex = 1
        Me.pageSynthese.Text = "Synthèse et lignes écartées"
        Me.pageSynthese.UseVisualStyleBackColor = True
        '
        'txtSynthese
        '
        ' Une police à chasse fixe : les colonnes de chiffres de la synthèse s'alignent.
        Me.txtSynthese.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtSynthese.Font = New System.Drawing.Font("Consolas", 9.75!)
        Me.txtSynthese.Location = New System.Drawing.Point(6, 6)
        Me.txtSynthese.Multiline = True
        Me.txtSynthese.Name = "txtSynthese"
        Me.txtSynthese.ReadOnly = True
        Me.txtSynthese.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtSynthese.Size = New System.Drawing.Size(956, 312)
        Me.txtSynthese.TabIndex = 0
        Me.txtSynthese.WordWrap = False
        '
        'lblStatut
        '
        Me.lblStatut.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblStatut.Location = New System.Drawing.Point(12, 588)
        Me.lblStatut.Name = "lblStatut"
        Me.lblStatut.Size = New System.Drawing.Size(660, 34)
        Me.lblStatut.TabIndex = 4
        Me.lblStatut.Text = "Chargez un rapport de règlement, puis cliquez sur Calculer."
        '
        'btnCopier
        '
        Me.btnCopier.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCopier.Location = New System.Drawing.Point(688, 586)
        Me.btnCopier.Name = "btnCopier"
        Me.btnCopier.Size = New System.Drawing.Size(150, 32)
        Me.btnCopier.TabIndex = 5
        Me.btnCopier.Text = "Copier la synthèse"
        Me.btnCopier.UseVisualStyleBackColor = True
        '
        'btnFermer
        '
        Me.btnFermer.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFermer.Location = New System.Drawing.Point(848, 586)
        Me.btnFermer.Name = "btnFermer"
        Me.btnFermer.Size = New System.Drawing.Size(140, 32)
        Me.btnFermer.TabIndex = 6
        Me.btnFermer.Text = "Fermer"
        Me.btnFermer.UseVisualStyleBackColor = True
        '
        'ofdRapport
        '
        Me.ofdRapport.Filter = "Rapports Western Union (*.zip;*.txt)|*.zip;*.txt|Archives ZIP (*.zip)|*.zip|Fichiers texte (*.txt)|*.txt|Tous les fichiers (*.*)|*.*"
        Me.ofdRapport.Title = "Sélectionner le rapport de règlement Western Union (archive ZIP ou fichier texte)"
        '
        'FrmControleChange
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnFermer
        Me.ClientSize = New System.Drawing.Size(1000, 630)
        Me.Controls.Add(Me.btnFermer)
        Me.Controls.Add(Me.btnCopier)
        Me.Controls.Add(Me.lblStatut)
        Me.Controls.Add(Me.onglets)
        Me.Controls.Add(Me.grpSource)
        Me.Controls.Add(Me.lblIntro)
        Me.Controls.Add(Me.panelTitre)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        Me.MinimumSize = New System.Drawing.Size(900, 560)
        Me.Name = "FrmControleChange"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Contrôle des écarts de change"
        Me.panelTitre.ResumeLayout(False)
        Me.grpSource.ResumeLayout(False)
        Me.onglets.ResumeLayout(False)
        Me.pageDetail.ResumeLayout(False)
        CType(Me.dgvEcarts, System.ComponentModel.ISupportInitialize).EndInit()
        Me.panelFiltre.ResumeLayout(False)
        Me.pageSynthese.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()
    End Sub

    Friend WithEvents panelTitre As System.Windows.Forms.Panel
    Friend WithEvents lblTitre As System.Windows.Forms.Label
    Friend WithEvents lblIntro As System.Windows.Forms.Label
    Friend WithEvents grpSource As System.Windows.Forms.GroupBox
    Friend WithEvents btnCharger As System.Windows.Forms.Button
    Friend WithEvents lblFichier As System.Windows.Forms.Label
    Friend WithEvents lblParite As System.Windows.Forms.Label
    Friend WithEvents txtParite As System.Windows.Forms.TextBox
    Friend WithEvents chkEnAttente As System.Windows.Forms.CheckBox
    Friend WithEvents btnAfficher As System.Windows.Forms.Button
    Friend WithEvents onglets As System.Windows.Forms.TabControl
    Friend WithEvents pageDetail As System.Windows.Forms.TabPage
    Friend WithEvents dgvEcarts As System.Windows.Forms.DataGridView
    Friend WithEvents panelFiltre As System.Windows.Forms.Panel
    Friend WithEvents lblMtcn As System.Windows.Forms.Label
    Friend WithEvents txtMtcn As System.Windows.Forms.TextBox
    Friend WithEvents lblFiltre As System.Windows.Forms.Label
    Friend WithEvents pageSynthese As System.Windows.Forms.TabPage
    Friend WithEvents txtSynthese As System.Windows.Forms.TextBox
    Friend WithEvents lblStatut As System.Windows.Forms.Label
    Friend WithEvents btnCopier As System.Windows.Forms.Button
    Friend WithEvents btnFermer As System.Windows.Forms.Button
    Friend WithEvents ofdRapport As System.Windows.Forms.OpenFileDialog

End Class
