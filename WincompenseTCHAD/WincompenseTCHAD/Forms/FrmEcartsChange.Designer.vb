Option Strict On
Option Explicit On

Partial Class FrmEcartsChange
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
        Me.grpSource = New System.Windows.Forms.GroupBox()
        Me.btnCharger = New System.Windows.Forms.Button()
        Me.lblFichier = New System.Windows.Forms.Label()
        Me.lblDecoupage = New System.Windows.Forms.Label()
        Me.cboDecoupage = New System.Windows.Forms.ComboBox()
        Me.lblParite = New System.Windows.Forms.Label()
        Me.btnAfficher = New System.Windows.Forms.Button()
        Me.btnPiece = New System.Windows.Forms.Button()
        Me.onglets = New System.Windows.Forms.TabControl()
        Me.pagePiece = New System.Windows.Forms.TabPage()
        Me.dgvPiece = New System.Windows.Forms.DataGridView()
        Me.panelTotaux = New System.Windows.Forms.Panel()
        Me.lblTotaux = New System.Windows.Forms.Label()
        Me.pageDetail = New System.Windows.Forms.TabPage()
        Me.dgvEcarts = New System.Windows.Forms.DataGridView()
        Me.pageSynthese = New System.Windows.Forms.TabPage()
        Me.txtSynthese = New System.Windows.Forms.TextBox()
        Me.lblStatut = New System.Windows.Forms.Label()
        Me.btnExporter = New System.Windows.Forms.Button()
        Me.btnEnregistrer = New System.Windows.Forms.Button()
        Me.btnFermer = New System.Windows.Forms.Button()
        Me.ofdRapport = New System.Windows.Forms.OpenFileDialog()
        Me.sfdExport = New System.Windows.Forms.SaveFileDialog()
        Me.panelTitre.SuspendLayout()
        Me.grpSource.SuspendLayout()
        Me.onglets.SuspendLayout()
        Me.pagePiece.SuspendLayout()
        CType(Me.dgvPiece, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.panelTotaux.SuspendLayout()
        Me.pageDetail.SuspendLayout()
        CType(Me.dgvEcarts, System.ComponentModel.ISupportInitialize).BeginInit()
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
        Me.panelTitre.Size = New System.Drawing.Size(1000, 48)
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
        Me.lblTitre.Size = New System.Drawing.Size(1000, 48)
        Me.lblTitre.TabIndex = 0
        Me.lblTitre.Text = "Écarts de change — pièce comptable"
        Me.lblTitre.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'grpSource
        '
        Me.grpSource.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grpSource.Controls.Add(Me.btnCharger)
        Me.grpSource.Controls.Add(Me.lblFichier)
        Me.grpSource.Controls.Add(Me.lblDecoupage)
        Me.grpSource.Controls.Add(Me.cboDecoupage)
        Me.grpSource.Controls.Add(Me.lblParite)
        Me.grpSource.Controls.Add(Me.btnAfficher)
        Me.grpSource.Controls.Add(Me.btnPiece)
        Me.grpSource.Location = New System.Drawing.Point(12, 66)
        Me.grpSource.Name = "grpSource"
        Me.grpSource.Size = New System.Drawing.Size(1000, 108)
        Me.grpSource.TabIndex = 1
        Me.grpSource.TabStop = False
        Me.grpSource.Text = "Rapport de règlement, découpage des pièces et parité"
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
        Me.lblFichier.Size = New System.Drawing.Size(812, 18)
        Me.lblFichier.TabIndex = 1
        Me.lblFichier.Text = "Aucun rapport chargé."
        '
        'lblDecoupage
        '
        Me.lblDecoupage.Location = New System.Drawing.Point(12, 70)
        Me.lblDecoupage.Name = "lblDecoupage"
        Me.lblDecoupage.Size = New System.Drawing.Size(80, 18)
        Me.lblDecoupage.TabIndex = 2
        Me.lblDecoupage.Text = "Découpage"
        '
        'cboDecoupage
        '
        Me.cboDecoupage.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboDecoupage.Location = New System.Drawing.Point(98, 67)
        Me.cboDecoupage.Name = "cboDecoupage"
        Me.cboDecoupage.Size = New System.Drawing.Size(330, 24)
        Me.cboDecoupage.TabIndex = 3
        '
        'lblParite
        '
        Me.lblParite.ForeColor = System.Drawing.SystemColors.GrayText
        Me.lblParite.Location = New System.Drawing.Point(444, 70)
        Me.lblParite.Name = "lblParite"
        Me.lblParite.Size = New System.Drawing.Size(230, 18)
        Me.lblParite.TabIndex = 4
        Me.lblParite.Text = "Parité : lue à l'ouverture du rapport."
        '
        'btnAfficher
        '
        Me.btnAfficher.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnAfficher.Location = New System.Drawing.Point(686, 64)
        Me.btnAfficher.Name = "btnAfficher"
        Me.btnAfficher.Size = New System.Drawing.Size(130, 30)
        Me.btnAfficher.TabIndex = 5
        Me.btnAfficher.Text = "Calculer"
        Me.btnAfficher.UseVisualStyleBackColor = True
        '
        'btnPiece
        '
        Me.btnPiece.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnPiece.Location = New System.Drawing.Point(826, 64)
        Me.btnPiece.Name = "btnPiece"
        Me.btnPiece.Size = New System.Drawing.Size(160, 30)
        Me.btnPiece.TabIndex = 6
        Me.btnPiece.Text = "Produire la pièce"
        Me.btnPiece.UseVisualStyleBackColor = True
        '
        'onglets
        '
        Me.onglets.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.onglets.Controls.Add(Me.pagePiece)
        Me.onglets.Controls.Add(Me.pageDetail)
        Me.onglets.Controls.Add(Me.pageSynthese)
        Me.onglets.Location = New System.Drawing.Point(12, 184)
        Me.onglets.Name = "onglets"
        Me.onglets.SelectedIndex = 0
        Me.onglets.Size = New System.Drawing.Size(1000, 388)
        Me.onglets.TabIndex = 2
        '
        'pagePiece
        '
        Me.pagePiece.Controls.Add(Me.dgvPiece)
        Me.pagePiece.Controls.Add(Me.panelTotaux)
        Me.pagePiece.Location = New System.Drawing.Point(4, 22)
        Me.pagePiece.Name = "pagePiece"
        Me.pagePiece.Padding = New System.Windows.Forms.Padding(6)
        Me.pagePiece.Size = New System.Drawing.Size(992, 362)
        Me.pagePiece.TabIndex = 0
        Me.pagePiece.Text = "Pièce comptable"
        Me.pagePiece.UseVisualStyleBackColor = True
        '
        'dgvPiece
        '
        Me.dgvPiece.AllowUserToAddRows = False
        Me.dgvPiece.AllowUserToDeleteRows = False
        Me.dgvPiece.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvPiece.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvPiece.Location = New System.Drawing.Point(6, 6)
        Me.dgvPiece.MultiSelect = False
        Me.dgvPiece.Name = "dgvPiece"
        Me.dgvPiece.ReadOnly = True
        Me.dgvPiece.RowHeadersWidth = 25
        Me.dgvPiece.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvPiece.Size = New System.Drawing.Size(980, 312)
        Me.dgvPiece.TabIndex = 1
        '
        'panelTotaux
        '
        Me.panelTotaux.Controls.Add(Me.lblTotaux)
        Me.panelTotaux.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.panelTotaux.Location = New System.Drawing.Point(6, 318)
        Me.panelTotaux.Name = "panelTotaux"
        Me.panelTotaux.Size = New System.Drawing.Size(980, 38)
        Me.panelTotaux.TabIndex = 0
        '
        'lblTotaux
        '
        Me.lblTotaux.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblTotaux.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold)
        Me.lblTotaux.Location = New System.Drawing.Point(0, 0)
        Me.lblTotaux.Name = "lblTotaux"
        Me.lblTotaux.Size = New System.Drawing.Size(980, 38)
        Me.lblTotaux.TabIndex = 0
        Me.lblTotaux.Text = "Aucune pièce produite."
        Me.lblTotaux.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'pageDetail
        '
        Me.pageDetail.Controls.Add(Me.dgvEcarts)
        Me.pageDetail.Location = New System.Drawing.Point(4, 22)
        Me.pageDetail.Name = "pageDetail"
        Me.pageDetail.Padding = New System.Windows.Forms.Padding(6)
        Me.pageDetail.Size = New System.Drawing.Size(992, 362)
        Me.pageDetail.TabIndex = 1
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
        Me.dgvEcarts.Size = New System.Drawing.Size(980, 350)
        Me.dgvEcarts.TabIndex = 0
        '
        'pageSynthese
        '
        Me.pageSynthese.Controls.Add(Me.txtSynthese)
        Me.pageSynthese.Location = New System.Drawing.Point(4, 22)
        Me.pageSynthese.Name = "pageSynthese"
        Me.pageSynthese.Padding = New System.Windows.Forms.Padding(6)
        Me.pageSynthese.Size = New System.Drawing.Size(992, 362)
        Me.pageSynthese.TabIndex = 2
        Me.pageSynthese.Text = "Synthèse"
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
        Me.txtSynthese.Size = New System.Drawing.Size(980, 350)
        Me.txtSynthese.TabIndex = 0
        Me.txtSynthese.WordWrap = False
        '
        'lblStatut
        '
        Me.lblStatut.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblStatut.Location = New System.Drawing.Point(12, 584)
        Me.lblStatut.Name = "lblStatut"
        Me.lblStatut.Size = New System.Drawing.Size(534, 48)
        Me.lblStatut.TabIndex = 3
        Me.lblStatut.Text = "Chargez un rapport de règlement, puis calculez."
        '
        'btnExporter
        '
        Me.btnExporter.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnExporter.Location = New System.Drawing.Point(562, 590)
        Me.btnExporter.Name = "btnExporter"
        Me.btnExporter.Size = New System.Drawing.Size(150, 32)
        Me.btnExporter.TabIndex = 4
        Me.btnExporter.Text = "Exporter"
        Me.btnExporter.UseVisualStyleBackColor = True
        '
        'btnEnregistrer
        '
        Me.btnEnregistrer.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnEnregistrer.Location = New System.Drawing.Point(722, 590)
        Me.btnEnregistrer.Name = "btnEnregistrer"
        Me.btnEnregistrer.Size = New System.Drawing.Size(170, 32)
        Me.btnEnregistrer.TabIndex = 5
        Me.btnEnregistrer.Text = "Conserver la pièce"
        Me.btnEnregistrer.UseVisualStyleBackColor = True
        '
        'btnFermer
        '
        Me.btnFermer.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFermer.Location = New System.Drawing.Point(902, 590)
        Me.btnFermer.Name = "btnFermer"
        Me.btnFermer.Size = New System.Drawing.Size(110, 32)
        Me.btnFermer.TabIndex = 6
        Me.btnFermer.Text = "Fermer"
        Me.btnFermer.UseVisualStyleBackColor = True
        '
        'ofdRapport
        '
        Me.ofdRapport.Filter = "Rapports Western Union (*.zip;*.txt)|*.zip;*.txt|Archives ZIP (*.zip)|*.zip|Fichiers texte (*.txt)|*.txt|Tous les fichiers (*.*)|*.*"
        Me.ofdRapport.Title = "Sélectionner le rapport de règlement Western Union (archive ZIP ou fichier texte)"
        '
        'sfdExport
        '
        Me.sfdExport.DefaultExt = "xlsx"
        Me.sfdExport.Filter = "Classeur Excel (*.xlsx)|*.xlsx"
        Me.sfdExport.Title = "Enregistrer la pièce de change"
        '
        'FrmEcartsChange
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnFermer
        Me.ClientSize = New System.Drawing.Size(1024, 640)
        Me.Controls.Add(Me.btnFermer)
        Me.Controls.Add(Me.btnEnregistrer)
        Me.Controls.Add(Me.btnExporter)
        Me.Controls.Add(Me.lblStatut)
        Me.Controls.Add(Me.onglets)
        Me.Controls.Add(Me.grpSource)
        Me.Controls.Add(Me.panelTitre)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        Me.MinimumSize = New System.Drawing.Size(920, 580)
        Me.Name = "FrmEcartsChange"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Écarts de change — pièce comptable"
        Me.panelTitre.ResumeLayout(False)
        Me.grpSource.ResumeLayout(False)
        Me.onglets.ResumeLayout(False)
        Me.pagePiece.ResumeLayout(False)
        CType(Me.dgvPiece, System.ComponentModel.ISupportInitialize).EndInit()
        Me.panelTotaux.ResumeLayout(False)
        Me.pageDetail.ResumeLayout(False)
        CType(Me.dgvEcarts, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pageSynthese.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()
    End Sub

    Friend WithEvents panelTitre As System.Windows.Forms.Panel
    Friend WithEvents lblTitre As System.Windows.Forms.Label
    Friend WithEvents grpSource As System.Windows.Forms.GroupBox
    Friend WithEvents btnCharger As System.Windows.Forms.Button
    Friend WithEvents lblFichier As System.Windows.Forms.Label
    Friend WithEvents lblDecoupage As System.Windows.Forms.Label
    Friend WithEvents cboDecoupage As System.Windows.Forms.ComboBox
    Friend WithEvents lblParite As System.Windows.Forms.Label
    Friend WithEvents btnAfficher As System.Windows.Forms.Button
    Friend WithEvents btnPiece As System.Windows.Forms.Button
    Friend WithEvents onglets As System.Windows.Forms.TabControl
    Friend WithEvents pagePiece As System.Windows.Forms.TabPage
    Friend WithEvents dgvPiece As System.Windows.Forms.DataGridView
    Friend WithEvents panelTotaux As System.Windows.Forms.Panel
    Friend WithEvents lblTotaux As System.Windows.Forms.Label
    Friend WithEvents pageDetail As System.Windows.Forms.TabPage
    Friend WithEvents dgvEcarts As System.Windows.Forms.DataGridView
    Friend WithEvents pageSynthese As System.Windows.Forms.TabPage
    Friend WithEvents txtSynthese As System.Windows.Forms.TextBox
    Friend WithEvents lblStatut As System.Windows.Forms.Label
    Friend WithEvents btnExporter As System.Windows.Forms.Button
    Friend WithEvents btnEnregistrer As System.Windows.Forms.Button
    Friend WithEvents btnFermer As System.Windows.Forms.Button
    Friend WithEvents ofdRapport As System.Windows.Forms.OpenFileDialog
    Friend WithEvents sfdExport As System.Windows.Forms.SaveFileDialog

End Class
