Option Strict On
Option Explicit On

Partial Class FrmParametrageFichier
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
        Me.ongletsParametrage = New System.Windows.Forms.TabControl()
        Me.pageExport = New System.Windows.Forms.TabPage()
        Me.lblExplicationExport = New System.Windows.Forms.Label()
        Me.dgvContenu = New System.Windows.Forms.DataGridView()
        Me.btnExporter = New System.Windows.Forms.Button()
        Me.lblResultatExport = New System.Windows.Forms.Label()
        Me.pageImport = New System.Windows.Forms.TabPage()
        Me.lblExplicationImport = New System.Windows.Forms.Label()
        Me.btnChoisir = New System.Windows.Forms.Button()
        Me.lblFichier = New System.Windows.Forms.Label()
        Me.lblProvenance = New System.Windows.Forms.Label()
        Me.dgvAnalyse = New System.Windows.Forms.DataGridView()
        Me.lblComptes = New System.Windows.Forms.Label()
        Me.dgvComptes = New System.Windows.Forms.DataGridView()
        Me.chkRemplacerComptes = New System.Windows.Forms.CheckBox()
        Me.btnCharger = New System.Windows.Forms.Button()
        Me.lblResultatImport = New System.Windows.Forms.Label()
        Me.btnFermer = New System.Windows.Forms.Button()
        Me.ofdParametrage = New System.Windows.Forms.OpenFileDialog()
        Me.sfdParametrage = New System.Windows.Forms.SaveFileDialog()
        Me.ongletsParametrage.SuspendLayout()
        Me.pageExport.SuspendLayout()
        Me.pageImport.SuspendLayout()
        CType(Me.dgvContenu, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvAnalyse, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvComptes, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'lblTitre
        '
        Me.lblTitre.BackColor = System.Drawing.Color.Black
        Me.lblTitre.Font = New System.Drawing.Font("Segoe UI", 14.0!, System.Drawing.FontStyle.Bold)
        Me.lblTitre.ForeColor = System.Drawing.Color.Yellow
        Me.lblTitre.Location = New System.Drawing.Point(12, 12)
        Me.lblTitre.Name = "lblTitre"
        Me.lblTitre.Size = New System.Drawing.Size(956, 34)
        Me.lblTitre.TabIndex = 0
        Me.lblTitre.Text = "   Paramétrage : fichier de secours"
        Me.lblTitre.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblSousTitre
        '
        Me.lblSousTitre.ForeColor = System.Drawing.SystemColors.GrayText
        Me.lblSousTitre.Location = New System.Drawing.Point(12, 50)
        Me.lblSousTitre.Name = "lblSousTitre"
        Me.lblSousTitre.Size = New System.Drawing.Size(956, 32)
        Me.lblSousTitre.TabIndex = 1
        Me.lblSousTitre.Text = "Ce fichier n'est pas une sauvegarde de la base — celle-ci se fait sur le serveur. " &
            "C'est de quoi remonter une installation neuve sans tout resaisir."
        '
        'ongletsParametrage
        '
        Me.ongletsParametrage.Anchor = CType(((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) Or System.Windows.Forms.AnchorStyles.Right)), System.Windows.Forms.AnchorStyles)
        Me.ongletsParametrage.Controls.Add(Me.pageExport)
        Me.ongletsParametrage.Controls.Add(Me.pageImport)
        Me.ongletsParametrage.Location = New System.Drawing.Point(12, 88)
        Me.ongletsParametrage.Name = "ongletsParametrage"
        Me.ongletsParametrage.SelectedIndex = 0
        Me.ongletsParametrage.Size = New System.Drawing.Size(956, 528)
        Me.ongletsParametrage.TabIndex = 2
        '
        'pageExport
        '
        Me.pageExport.Controls.Add(Me.lblExplicationExport)
        Me.pageExport.Controls.Add(Me.dgvContenu)
        Me.pageExport.Controls.Add(Me.btnExporter)
        Me.pageExport.Controls.Add(Me.lblResultatExport)
        Me.pageExport.Location = New System.Drawing.Point(4, 22)
        Me.pageExport.Name = "pageExport"
        Me.pageExport.Padding = New System.Windows.Forms.Padding(3)
        Me.pageExport.Size = New System.Drawing.Size(948, 502)
        Me.pageExport.TabIndex = 0
        Me.pageExport.Text = "Exporter le paramétrage"
        Me.pageExport.UseVisualStyleBackColor = True
        '
        'lblExplicationExport
        '
        Me.lblExplicationExport.Location = New System.Drawing.Point(14, 14)
        Me.lblExplicationExport.Name = "lblExplicationExport"
        Me.lblExplicationExport.Size = New System.Drawing.Size(920, 56)
        Me.lblExplicationExport.TabIndex = 0
        Me.lblExplicationExport.Text = "L'archive contiendra les comptes comptables, les groupes statistiques, les sous-ag" &
            "ents et les agences propres. Les utilisateurs y figurent pour mémoire : aucun mot" &
            " de passe n'en sort, pas même sous forme d'empreinte."
        '
        'dgvContenu
        '
        Me.dgvContenu.AllowUserToAddRows = False
        Me.dgvContenu.AllowUserToDeleteRows = False
        Me.dgvContenu.Anchor = CType(((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) Or System.Windows.Forms.AnchorStyles.Right)), System.Windows.Forms.AnchorStyles)
        Me.dgvContenu.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvContenu.Location = New System.Drawing.Point(14, 76)
        Me.dgvContenu.MultiSelect = False
        Me.dgvContenu.Name = "dgvContenu"
        Me.dgvContenu.ReadOnly = True
        Me.dgvContenu.RowHeadersWidth = 25
        Me.dgvContenu.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvContenu.Size = New System.Drawing.Size(920, 336)
        Me.dgvContenu.TabIndex = 1
        '
        'btnExporter
        '
        Me.btnExporter.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnExporter.Location = New System.Drawing.Point(14, 426)
        Me.btnExporter.Name = "btnExporter"
        Me.btnExporter.Size = New System.Drawing.Size(230, 34)
        Me.btnExporter.TabIndex = 2
        Me.btnExporter.Text = "Exporter le paramétrage…"
        Me.btnExporter.UseVisualStyleBackColor = True
        '
        'lblResultatExport
        '
        Me.lblResultatExport.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblResultatExport.ForeColor = System.Drawing.SystemColors.GrayText
        Me.lblResultatExport.Location = New System.Drawing.Point(254, 426)
        Me.lblResultatExport.Name = "lblResultatExport"
        Me.lblResultatExport.Size = New System.Drawing.Size(680, 62)
        Me.lblResultatExport.TabIndex = 3
        '
        'pageImport
        '
        Me.pageImport.Controls.Add(Me.lblExplicationImport)
        Me.pageImport.Controls.Add(Me.btnChoisir)
        Me.pageImport.Controls.Add(Me.lblFichier)
        Me.pageImport.Controls.Add(Me.lblProvenance)
        Me.pageImport.Controls.Add(Me.dgvAnalyse)
        Me.pageImport.Controls.Add(Me.lblComptes)
        Me.pageImport.Controls.Add(Me.dgvComptes)
        Me.pageImport.Controls.Add(Me.chkRemplacerComptes)
        Me.pageImport.Controls.Add(Me.btnCharger)
        Me.pageImport.Controls.Add(Me.lblResultatImport)
        Me.pageImport.Location = New System.Drawing.Point(4, 22)
        Me.pageImport.Name = "pageImport"
        Me.pageImport.Padding = New System.Windows.Forms.Padding(3)
        Me.pageImport.Size = New System.Drawing.Size(948, 502)
        Me.pageImport.TabIndex = 1
        Me.pageImport.Text = "Charger un paramétrage"
        Me.pageImport.UseVisualStyleBackColor = True
        '
        'lblExplicationImport
        '
        Me.lblExplicationImport.Location = New System.Drawing.Point(14, 12)
        Me.lblExplicationImport.Name = "lblExplicationImport"
        Me.lblExplicationImport.Size = New System.Drawing.Size(920, 44)
        Me.lblExplicationImport.TabIndex = 0
        Me.lblExplicationImport.Text = "Le chargement ne remplace JAMAIS une ligne déjà présente : il ne crée que ce qui m" &
            "anque. Seuls les comptes comptables font exception, et il faut alors le demander " &
            "explicitement, ci-dessous."
        '
        'btnChoisir
        '
        Me.btnChoisir.Location = New System.Drawing.Point(14, 60)
        Me.btnChoisir.Name = "btnChoisir"
        Me.btnChoisir.Size = New System.Drawing.Size(180, 30)
        Me.btnChoisir.TabIndex = 1
        Me.btnChoisir.Text = "Choisir l'archive…"
        Me.btnChoisir.UseVisualStyleBackColor = True
        '
        'lblFichier
        '
        Me.lblFichier.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblFichier.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold)
        Me.lblFichier.Location = New System.Drawing.Point(204, 60)
        Me.lblFichier.Name = "lblFichier"
        Me.lblFichier.Size = New System.Drawing.Size(730, 18)
        Me.lblFichier.TabIndex = 2
        Me.lblFichier.Text = "Aucune archive choisie."
        '
        'lblProvenance
        '
        Me.lblProvenance.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblProvenance.ForeColor = System.Drawing.SystemColors.GrayText
        Me.lblProvenance.Location = New System.Drawing.Point(204, 78)
        Me.lblProvenance.Name = "lblProvenance"
        Me.lblProvenance.Size = New System.Drawing.Size(730, 18)
        Me.lblProvenance.TabIndex = 3
        '
        'dgvAnalyse
        '
        Me.dgvAnalyse.AllowUserToAddRows = False
        Me.dgvAnalyse.AllowUserToDeleteRows = False
        Me.dgvAnalyse.Anchor = CType(((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) Or System.Windows.Forms.AnchorStyles.Right)), System.Windows.Forms.AnchorStyles)
        Me.dgvAnalyse.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvAnalyse.Location = New System.Drawing.Point(14, 104)
        Me.dgvAnalyse.MultiSelect = False
        Me.dgvAnalyse.Name = "dgvAnalyse"
        Me.dgvAnalyse.ReadOnly = True
        Me.dgvAnalyse.RowHeadersWidth = 25
        Me.dgvAnalyse.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvAnalyse.Size = New System.Drawing.Size(560, 348)
        Me.dgvAnalyse.TabIndex = 4
        '
        'lblComptes
        '
        Me.lblComptes.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblComptes.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold)
        Me.lblComptes.Location = New System.Drawing.Point(586, 104)
        Me.lblComptes.Name = "lblComptes"
        Me.lblComptes.Size = New System.Drawing.Size(348, 18)
        Me.lblComptes.TabIndex = 5
        Me.lblComptes.Text = "Comptes comptables"
        '
        'dgvComptes
        '
        Me.dgvComptes.AllowUserToAddRows = False
        Me.dgvComptes.AllowUserToDeleteRows = False
        Me.dgvComptes.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Right)), System.Windows.Forms.AnchorStyles)
        Me.dgvComptes.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvComptes.Location = New System.Drawing.Point(586, 126)
        Me.dgvComptes.MultiSelect = False
        Me.dgvComptes.Name = "dgvComptes"
        Me.dgvComptes.ReadOnly = True
        Me.dgvComptes.RowHeadersWidth = 25
        Me.dgvComptes.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvComptes.Size = New System.Drawing.Size(348, 296)
        Me.dgvComptes.TabIndex = 6
        '
        'chkRemplacerComptes
        '
        Me.chkRemplacerComptes.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.chkRemplacerComptes.Location = New System.Drawing.Point(586, 428)
        Me.chkRemplacerComptes.Name = "chkRemplacerComptes"
        Me.chkRemplacerComptes.Size = New System.Drawing.Size(348, 44)
        Me.chkRemplacerComptes.TabIndex = 7
        Me.chkRemplacerComptes.Text = "Remplacer les comptes comptables de la base par ceux du fichier"
        Me.chkRemplacerComptes.UseVisualStyleBackColor = True
        '
        'btnCharger
        '
        Me.btnCharger.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnCharger.Location = New System.Drawing.Point(14, 462)
        Me.btnCharger.Name = "btnCharger"
        Me.btnCharger.Size = New System.Drawing.Size(230, 34)
        Me.btnCharger.TabIndex = 8
        Me.btnCharger.Text = "Charger le paramétrage"
        Me.btnCharger.UseVisualStyleBackColor = True
        '
        'lblResultatImport
        '
        Me.lblResultatImport.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblResultatImport.Location = New System.Drawing.Point(254, 458)
        Me.lblResultatImport.Name = "lblResultatImport"
        Me.lblResultatImport.Size = New System.Drawing.Size(320, 40)
        Me.lblResultatImport.TabIndex = 9
        '
        'btnFermer
        '
        Me.btnFermer.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFermer.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnFermer.Location = New System.Drawing.Point(876, 624)
        Me.btnFermer.Name = "btnFermer"
        Me.btnFermer.Size = New System.Drawing.Size(92, 32)
        Me.btnFermer.TabIndex = 3
        Me.btnFermer.Text = "Fermer"
        Me.btnFermer.UseVisualStyleBackColor = True
        '
        'ofdParametrage
        '
        Me.ofdParametrage.Filter = "Archive de paramétrage (*.zip)|*.zip"
        Me.ofdParametrage.Title = "Choisir l'archive de paramétrage"
        '
        'sfdParametrage
        '
        Me.sfdParametrage.Filter = "Archive de paramétrage (*.zip)|*.zip"
        Me.sfdParametrage.Title = "Enregistrer le paramétrage"
        '
        'FrmParametrageFichier
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnFermer
        Me.ClientSize = New System.Drawing.Size(980, 668)
        Me.Controls.Add(Me.btnFermer)
        Me.Controls.Add(Me.ongletsParametrage)
        Me.Controls.Add(Me.lblSousTitre)
        Me.Controls.Add(Me.lblTitre)
        Me.MinimumSize = New System.Drawing.Size(880, 600)
        Me.Name = "FrmParametrageFichier"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Paramétrage : fichier de secours"
        Me.ongletsParametrage.ResumeLayout(False)
        Me.pageExport.ResumeLayout(False)
        Me.pageImport.ResumeLayout(False)
        CType(Me.dgvContenu, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvAnalyse, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvComptes, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents lblTitre As System.Windows.Forms.Label
    Friend WithEvents lblSousTitre As System.Windows.Forms.Label
    Friend WithEvents ongletsParametrage As System.Windows.Forms.TabControl
    Friend WithEvents pageExport As System.Windows.Forms.TabPage
    Friend WithEvents lblExplicationExport As System.Windows.Forms.Label
    Friend WithEvents dgvContenu As System.Windows.Forms.DataGridView
    Friend WithEvents btnExporter As System.Windows.Forms.Button
    Friend WithEvents lblResultatExport As System.Windows.Forms.Label
    Friend WithEvents pageImport As System.Windows.Forms.TabPage
    Friend WithEvents lblExplicationImport As System.Windows.Forms.Label
    Friend WithEvents btnChoisir As System.Windows.Forms.Button
    Friend WithEvents lblFichier As System.Windows.Forms.Label
    Friend WithEvents lblProvenance As System.Windows.Forms.Label
    Friend WithEvents dgvAnalyse As System.Windows.Forms.DataGridView
    Friend WithEvents lblComptes As System.Windows.Forms.Label
    Friend WithEvents dgvComptes As System.Windows.Forms.DataGridView
    Friend WithEvents chkRemplacerComptes As System.Windows.Forms.CheckBox
    Friend WithEvents btnCharger As System.Windows.Forms.Button
    Friend WithEvents lblResultatImport As System.Windows.Forms.Label
    Friend WithEvents btnFermer As System.Windows.Forms.Button
    Friend WithEvents ofdParametrage As System.Windows.Forms.OpenFileDialog
    Friend WithEvents sfdParametrage As System.Windows.Forms.SaveFileDialog

End Class
