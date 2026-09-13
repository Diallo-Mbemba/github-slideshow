Option Strict On
Option Explicit On

Partial Class FrmRapportActivite
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
        Me.lblDu = New System.Windows.Forms.Label()
        Me.lblAu = New System.Windows.Forms.Label()
        Me.lblGroupe = New System.Windows.Forms.Label()
        Me.lblDisponible = New System.Windows.Forms.Label()
        Me.lblStatut = New System.Windows.Forms.Label()
        Me.dtpDebut = New System.Windows.Forms.DateTimePicker()
        Me.dtpFin = New System.Windows.Forms.DateTimePicker()
        Me.cboGroupe = New System.Windows.Forms.ComboBox()
        Me.btnAfficher = New System.Windows.Forms.Button()
        Me.btnExporter = New System.Windows.Forms.Button()
        Me.btnFermer = New System.Windows.Forms.Button()
        Me.tabRapport = New System.Windows.Forms.TabControl()
        Me.tabSynthese = New System.Windows.Forms.TabPage()
        Me.dgvSynthese = New System.Windows.Forms.DataGridView()
        Me.tabParJour = New System.Windows.Forms.TabPage()
        Me.dgvParJour = New System.Windows.Forms.DataGridView()
        Me.tabParPdv = New System.Windows.Forms.TabPage()
        Me.dgvParPdv = New System.Windows.Forms.DataGridView()
        Me.tabParGroupe = New System.Windows.Forms.TabPage()
        Me.dgvParGroupe = New System.Windows.Forms.DataGridView()
        Me.tabCommissions = New System.Windows.Forms.TabPage()
        Me.tabMtcn = New System.Windows.Forms.TabPage()
        Me.dgvCommissions = New System.Windows.Forms.DataGridView()
        Me.dgvMtcn = New System.Windows.Forms.DataGridView()
        Me.sfdExport = New System.Windows.Forms.SaveFileDialog()
        Me.tabRapport.SuspendLayout()
        Me.tabSynthese.SuspendLayout()
        CType(Me.dgvSynthese, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tabParJour.SuspendLayout()
        CType(Me.dgvParJour, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tabParPdv.SuspendLayout()
        CType(Me.dgvParPdv, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tabParGroupe.SuspendLayout()
        CType(Me.dgvParGroupe, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tabCommissions.SuspendLayout()
        Me.tabMtcn.SuspendLayout()
        CType(Me.dgvCommissions, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvMtcn, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'lblDu
        '
        Me.lblDu.Location = New System.Drawing.Point(12, 16)
        Me.lblDu.Name = "lblDu"
        Me.lblDu.Size = New System.Drawing.Size(70, 20)
        Me.lblDu.TabIndex = 0
        Me.lblDu.Text = "Période du"
        '
        'dtpDebut
        '
        Me.dtpDebut.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtpDebut.Location = New System.Drawing.Point(85, 12)
        Me.dtpDebut.Name = "dtpDebut"
        Me.dtpDebut.Size = New System.Drawing.Size(120, 22)
        Me.dtpDebut.TabIndex = 1
        '
        'lblAu
        '
        Me.lblAu.Location = New System.Drawing.Point(215, 16)
        Me.lblAu.Name = "lblAu"
        Me.lblAu.Size = New System.Drawing.Size(25, 20)
        Me.lblAu.TabIndex = 2
        Me.lblAu.Text = "au"
        '
        'dtpFin
        '
        Me.dtpFin.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtpFin.Location = New System.Drawing.Point(243, 12)
        Me.dtpFin.Name = "dtpFin"
        Me.dtpFin.Size = New System.Drawing.Size(120, 22)
        Me.dtpFin.TabIndex = 3
        '
        'btnAfficher
        '
        Me.btnAfficher.Location = New System.Drawing.Point(375, 11)
        Me.btnAfficher.Name = "btnAfficher"
        Me.btnAfficher.Size = New System.Drawing.Size(130, 26)
        Me.btnAfficher.TabIndex = 4
        Me.btnAfficher.Text = "Afficher"
        Me.btnAfficher.UseVisualStyleBackColor = True
        '
        'lblDisponible
        '
        Me.lblDisponible.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblDisponible.ForeColor = System.Drawing.SystemColors.GrayText
        Me.lblDisponible.Location = New System.Drawing.Point(588, 16)
        Me.lblDisponible.Name = "lblDisponible"
        Me.lblDisponible.Size = New System.Drawing.Size(440, 20)
        Me.lblDisponible.TabIndex = 5
        Me.lblDisponible.Text = ""
        Me.lblDisponible.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblGroupe
        '
        Me.lblGroupe.Location = New System.Drawing.Point(12, 50)
        Me.lblGroupe.Name = "lblGroupe"
        Me.lblGroupe.Size = New System.Drawing.Size(120, 20)
        Me.lblGroupe.TabIndex = 6
        Me.lblGroupe.Text = "Groupe statistique"
        '
        'cboGroupe
        '
        Me.cboGroupe.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboGroupe.FormattingEnabled = True
        Me.cboGroupe.Location = New System.Drawing.Point(135, 46)
        Me.cboGroupe.Name = "cboGroupe"
        Me.cboGroupe.Size = New System.Drawing.Size(300, 24)
        Me.cboGroupe.TabIndex = 7
        '
        'tabRapport
        '
        Me.tabRapport.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tabRapport.Controls.Add(Me.tabSynthese)
        Me.tabRapport.Controls.Add(Me.tabParJour)
        Me.tabRapport.Controls.Add(Me.tabParPdv)
        Me.tabRapport.Controls.Add(Me.tabParGroupe)
        Me.tabRapport.Controls.Add(Me.tabCommissions)
        Me.tabRapport.Controls.Add(Me.tabMtcn)
        Me.tabRapport.Location = New System.Drawing.Point(12, 80)
        Me.tabRapport.Name = "tabRapport"
        Me.tabRapport.SelectedIndex = 0
        Me.tabRapport.Size = New System.Drawing.Size(1016, 512)
        Me.tabRapport.TabIndex = 8
        '
        'tabSynthese
        '
        Me.tabSynthese.Controls.Add(Me.dgvSynthese)
        Me.tabSynthese.Location = New System.Drawing.Point(4, 22)
        Me.tabSynthese.Name = "tabSynthese"
        Me.tabSynthese.Padding = New System.Windows.Forms.Padding(3)
        Me.tabSynthese.Size = New System.Drawing.Size(1008, 486)
        Me.tabSynthese.TabIndex = 0
        Me.tabSynthese.Text = "Synthèse"
        Me.tabSynthese.UseVisualStyleBackColor = True
        '
        'dgvSynthese
        '
        Me.dgvSynthese.AllowUserToAddRows = False
        Me.dgvSynthese.AllowUserToDeleteRows = False
        Me.dgvSynthese.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvSynthese.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvSynthese.Location = New System.Drawing.Point(3, 3)
        Me.dgvSynthese.MultiSelect = False
        Me.dgvSynthese.Name = "dgvSynthese"
        Me.dgvSynthese.ReadOnly = True
        Me.dgvSynthese.RowHeadersWidth = 25
        Me.dgvSynthese.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvSynthese.Size = New System.Drawing.Size(1002, 480)
        Me.dgvSynthese.TabIndex = 0
        '
        'tabParJour
        '
        Me.tabParJour.Controls.Add(Me.dgvParJour)
        Me.tabParJour.Location = New System.Drawing.Point(4, 22)
        Me.tabParJour.Name = "tabParJour"
        Me.tabParJour.Padding = New System.Windows.Forms.Padding(3)
        Me.tabParJour.Size = New System.Drawing.Size(1008, 486)
        Me.tabParJour.TabIndex = 1
        Me.tabParJour.Text = "Jour par jour"
        Me.tabParJour.UseVisualStyleBackColor = True
        '
        'dgvParJour
        '
        Me.dgvParJour.AllowUserToAddRows = False
        Me.dgvParJour.AllowUserToDeleteRows = False
        Me.dgvParJour.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvParJour.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvParJour.Location = New System.Drawing.Point(3, 3)
        Me.dgvParJour.MultiSelect = False
        Me.dgvParJour.Name = "dgvParJour"
        Me.dgvParJour.ReadOnly = True
        Me.dgvParJour.RowHeadersWidth = 25
        Me.dgvParJour.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvParJour.Size = New System.Drawing.Size(1002, 480)
        Me.dgvParJour.TabIndex = 0
        '
        'tabParPdv
        '
        Me.tabParPdv.Controls.Add(Me.dgvParPdv)
        Me.tabParPdv.Location = New System.Drawing.Point(4, 22)
        Me.tabParPdv.Name = "tabParPdv"
        Me.tabParPdv.Padding = New System.Windows.Forms.Padding(3)
        Me.tabParPdv.Size = New System.Drawing.Size(1008, 486)
        Me.tabParPdv.TabIndex = 2
        Me.tabParPdv.Text = "Par point de vente"
        Me.tabParPdv.UseVisualStyleBackColor = True
        '
        'dgvParPdv
        '
        Me.dgvParPdv.AllowUserToAddRows = False
        Me.dgvParPdv.AllowUserToDeleteRows = False
        Me.dgvParPdv.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvParPdv.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvParPdv.Location = New System.Drawing.Point(3, 3)
        Me.dgvParPdv.MultiSelect = False
        Me.dgvParPdv.Name = "dgvParPdv"
        Me.dgvParPdv.ReadOnly = True
        Me.dgvParPdv.RowHeadersWidth = 25
        Me.dgvParPdv.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvParPdv.Size = New System.Drawing.Size(1002, 480)
        Me.dgvParPdv.TabIndex = 0
        '
        'tabParGroupe
        '
        Me.tabParGroupe.Controls.Add(Me.dgvParGroupe)
        Me.tabParGroupe.Location = New System.Drawing.Point(4, 22)
        Me.tabParGroupe.Name = "tabParGroupe"
        Me.tabParGroupe.Padding = New System.Windows.Forms.Padding(3)
        Me.tabParGroupe.Size = New System.Drawing.Size(1008, 486)
        Me.tabParGroupe.TabIndex = 3
        Me.tabParGroupe.Text = "Par groupe statistique"
        Me.tabParGroupe.UseVisualStyleBackColor = True
        '
        'dgvParGroupe
        '
        Me.dgvParGroupe.AllowUserToAddRows = False
        Me.dgvParGroupe.AllowUserToDeleteRows = False
        Me.dgvParGroupe.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvParGroupe.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvParGroupe.Location = New System.Drawing.Point(3, 3)
        Me.dgvParGroupe.MultiSelect = False
        Me.dgvParGroupe.Name = "dgvParGroupe"
        Me.dgvParGroupe.ReadOnly = True
        Me.dgvParGroupe.RowHeadersWidth = 25
        Me.dgvParGroupe.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvParGroupe.Size = New System.Drawing.Size(1002, 480)
        Me.dgvParGroupe.TabIndex = 0
        '
        'tabCommissions
        '
        Me.tabCommissions.Controls.Add(Me.dgvCommissions)
        Me.tabCommissions.Location = New System.Drawing.Point(4, 22)
        Me.tabCommissions.Name = "tabCommissions"
        Me.tabCommissions.Padding = New System.Windows.Forms.Padding(3)
        Me.tabCommissions.Size = New System.Drawing.Size(1008, 486)
        Me.tabCommissions.TabIndex = 4
        Me.tabCommissions.Text = "Évolution des commissions"
        Me.tabCommissions.UseVisualStyleBackColor = True
        '
        'dgvCommissions
        '
        Me.dgvCommissions.AllowUserToAddRows = False
        Me.dgvCommissions.AllowUserToDeleteRows = False
        Me.dgvCommissions.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvCommissions.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvCommissions.Location = New System.Drawing.Point(3, 3)
        Me.dgvCommissions.MultiSelect = False
        Me.dgvCommissions.Name = "dgvCommissions"
        Me.dgvCommissions.ReadOnly = True
        Me.dgvCommissions.RowHeadersWidth = 25
        Me.dgvCommissions.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvCommissions.Size = New System.Drawing.Size(1002, 480)
        Me.dgvCommissions.TabIndex = 0
        '
        'btnExporter
        '
        Me.btnExporter.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnExporter.Location = New System.Drawing.Point(12, 602)
        Me.btnExporter.Name = "btnExporter"
        Me.btnExporter.Size = New System.Drawing.Size(240, 32)
        Me.btnExporter.TabIndex = 9
        Me.btnExporter.Text = "Exporter en PDF..."
        Me.btnExporter.UseVisualStyleBackColor = True
        '
        'btnFermer
        '
        Me.btnFermer.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFermer.Location = New System.Drawing.Point(908, 602)
        Me.btnFermer.Name = "btnFermer"
        Me.btnFermer.Size = New System.Drawing.Size(120, 32)
        Me.btnFermer.TabIndex = 9
        Me.btnFermer.Text = "Fermer"
        Me.btnFermer.UseVisualStyleBackColor = True
        '
        'sfdExport
        '
        Me.sfdExport.DefaultExt = "pdf"
        Me.sfdExport.Filter = "Document PDF (*.pdf)|*.pdf"
        Me.sfdExport.Title = "Exporter le rapport d'activité"
        '
        'lblStatut
        '
        Me.lblStatut.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblStatut.ForeColor = System.Drawing.SystemColors.GrayText
        Me.lblStatut.Location = New System.Drawing.Point(262, 610)
        Me.lblStatut.Name = "lblStatut"
        Me.lblStatut.Size = New System.Drawing.Size(628, 20)
        Me.lblStatut.TabIndex = 11
        Me.lblStatut.Text = ""
        '
        'tabMtcn
        '
        Me.tabMtcn.Controls.Add(Me.dgvMtcn)
        Me.tabMtcn.Location = New System.Drawing.Point(4, 22)
        Me.tabMtcn.Name = "tabMtcn"
        Me.tabMtcn.Padding = New System.Windows.Forms.Padding(3)
        Me.tabMtcn.Size = New System.Drawing.Size(1008, 486)
        Me.tabMtcn.TabIndex = 5
        Me.tabMtcn.Text = "Transactions (MTCN)"
        Me.tabMtcn.UseVisualStyleBackColor = True
        '
        'dgvMtcn
        '
        Me.dgvMtcn.AllowUserToAddRows = False
        Me.dgvMtcn.AllowUserToDeleteRows = False
        Me.dgvMtcn.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvMtcn.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvMtcn.Location = New System.Drawing.Point(3, 3)
        Me.dgvMtcn.MultiSelect = False
        Me.dgvMtcn.Name = "dgvMtcn"
        Me.dgvMtcn.ReadOnly = True
        Me.dgvMtcn.RowHeadersWidth = 25
        Me.dgvMtcn.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvMtcn.Size = New System.Drawing.Size(1002, 480)
        Me.dgvMtcn.TabIndex = 0
        '
        'FrmRapportActivite
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnFermer
        Me.ClientSize = New System.Drawing.Size(1040, 680)
        Me.Controls.Add(Me.lblDu)
        Me.Controls.Add(Me.dtpDebut)
        Me.Controls.Add(Me.lblAu)
        Me.Controls.Add(Me.dtpFin)
        Me.Controls.Add(Me.btnAfficher)
        Me.Controls.Add(Me.lblDisponible)
        Me.Controls.Add(Me.lblGroupe)
        Me.Controls.Add(Me.cboGroupe)
        Me.Controls.Add(Me.tabRapport)
        Me.Controls.Add(Me.btnExporter)
        Me.Controls.Add(Me.btnFermer)
        Me.Controls.Add(Me.lblStatut)
        Me.MinimumSize = New System.Drawing.Size(840, 520)
        Me.Name = "FrmRapportActivite"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Rapport d'activité Western Union"
        Me.tabRapport.ResumeLayout(False)
        Me.tabSynthese.ResumeLayout(False)
        CType(Me.dgvSynthese, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tabParJour.ResumeLayout(False)
        CType(Me.dgvParJour, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tabParPdv.ResumeLayout(False)
        CType(Me.dgvParPdv, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tabParGroupe.ResumeLayout(False)
        CType(Me.dgvParGroupe, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tabCommissions.ResumeLayout(False)
        Me.tabMtcn.ResumeLayout(False)
        CType(Me.dgvCommissions, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvMtcn, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
    End Sub

    Friend WithEvents lblDu As System.Windows.Forms.Label
    Friend WithEvents lblAu As System.Windows.Forms.Label
    Friend WithEvents lblGroupe As System.Windows.Forms.Label
    Friend WithEvents lblDisponible As System.Windows.Forms.Label
    Friend WithEvents lblStatut As System.Windows.Forms.Label
    Friend WithEvents dtpDebut As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtpFin As System.Windows.Forms.DateTimePicker
    Friend WithEvents cboGroupe As System.Windows.Forms.ComboBox
    Friend WithEvents btnAfficher As System.Windows.Forms.Button
    Friend WithEvents btnExporter As System.Windows.Forms.Button
    Friend WithEvents btnFermer As System.Windows.Forms.Button
    Friend WithEvents tabRapport As System.Windows.Forms.TabControl
    Friend WithEvents tabSynthese As System.Windows.Forms.TabPage
    Friend WithEvents dgvSynthese As System.Windows.Forms.DataGridView
    Friend WithEvents tabParJour As System.Windows.Forms.TabPage
    Friend WithEvents dgvParJour As System.Windows.Forms.DataGridView
    Friend WithEvents tabParPdv As System.Windows.Forms.TabPage
    Friend WithEvents dgvParPdv As System.Windows.Forms.DataGridView
    Friend WithEvents tabParGroupe As System.Windows.Forms.TabPage
    Friend WithEvents dgvParGroupe As System.Windows.Forms.DataGridView
    Friend WithEvents tabCommissions As System.Windows.Forms.TabPage
    Friend WithEvents tabMtcn As System.Windows.Forms.TabPage
    Friend WithEvents dgvCommissions As System.Windows.Forms.DataGridView
    Friend WithEvents dgvMtcn As System.Windows.Forms.DataGridView
    Friend WithEvents sfdExport As System.Windows.Forms.SaveFileDialog

End Class
