Option Strict On
Option Explicit On

Partial Class FrmCommissionsBanque
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
        Me.lblDu = New System.Windows.Forms.Label()
        Me.dtpDebut = New System.Windows.Forms.DateTimePicker()
        Me.lblAu = New System.Windows.Forms.Label()
        Me.dtpFin = New System.Windows.Forms.DateTimePicker()
        Me.btnAfficher = New System.Windows.Forms.Button()
        Me.lblDisponible = New System.Windows.Forms.Label()
        Me.btnMois = New System.Windows.Forms.Button()
        Me.chkASigner = New System.Windows.Forms.CheckBox()
        Me.lblAvertissement = New System.Windows.Forms.Label()
        Me.lblSynthese = New System.Windows.Forms.Label()
        Me.dgvSynthese = New System.Windows.Forms.DataGridView()
        Me.lblComptes = New System.Windows.Forms.Label()
        Me.dgvComptes = New System.Windows.Forms.DataGridView()
        Me.lblParJour = New System.Windows.Forms.Label()
        Me.dgvParJour = New System.Windows.Forms.DataGridView()
        Me.lblControle = New System.Windows.Forms.Label()
        Me.btnExporter = New System.Windows.Forms.Button()
        Me.lblStatut = New System.Windows.Forms.Label()
        Me.btnFermer = New System.Windows.Forms.Button()
        Me.sfdEtat = New System.Windows.Forms.SaveFileDialog()
        CType(Me.dgvSynthese, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvComptes, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvParJour, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'lblTitre
        '
        Me.lblTitre.BackColor = System.Drawing.Color.Black
        Me.lblTitre.Font = New System.Drawing.Font("Segoe UI", 14.0!, System.Drawing.FontStyle.Bold)
        Me.lblTitre.ForeColor = System.Drawing.Color.Yellow
        Me.lblTitre.Location = New System.Drawing.Point(12, 12)
        Me.lblTitre.Name = "lblTitre"
        Me.lblTitre.Size = New System.Drawing.Size(1136, 34)
        Me.lblTitre.TabIndex = 0
        Me.lblTitre.Text = "   Commissions encaissées par la banque"
        Me.lblTitre.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblSousTitre
        '
        Me.lblSousTitre.ForeColor = System.Drawing.SystemColors.GrayText
        Me.lblSousTitre.Location = New System.Drawing.Point(12, 50)
        Me.lblSousTitre.Name = "lblSousTitre"
        Me.lblSousTitre.Size = New System.Drawing.Size(1136, 32)
        Me.lblSousTitre.TabIndex = 1
        Me.lblSousTitre.Text = "Ce que la banque garde réellement : sur un sous-agent, la part non rétrocédée ; su" &
            "r une agence propre, la totalité. L'onglet « Évolution des commissions » des rapp" &
            "orts montre, lui, la commission générée — rétrocession comprise."
        '
        'lblDu
        '
        Me.lblDu.Location = New System.Drawing.Point(12, 92)
        Me.lblDu.Name = "lblDu"
        Me.lblDu.Size = New System.Drawing.Size(70, 20)
        Me.lblDu.TabIndex = 2
        Me.lblDu.Text = "Période du"
        '
        'dtpDebut
        '
        Me.dtpDebut.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtpDebut.Location = New System.Drawing.Point(84, 88)
        Me.dtpDebut.Name = "dtpDebut"
        Me.dtpDebut.Size = New System.Drawing.Size(110, 20)
        Me.dtpDebut.TabIndex = 3
        '
        'lblAu
        '
        Me.lblAu.Location = New System.Drawing.Point(202, 92)
        Me.lblAu.Name = "lblAu"
        Me.lblAu.Size = New System.Drawing.Size(28, 20)
        Me.lblAu.TabIndex = 4
        Me.lblAu.Text = "au"
        '
        'dtpFin
        '
        Me.dtpFin.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtpFin.Location = New System.Drawing.Point(234, 88)
        Me.dtpFin.Name = "dtpFin"
        Me.dtpFin.Size = New System.Drawing.Size(110, 20)
        Me.dtpFin.TabIndex = 5
        '
        'btnAfficher
        '
        Me.btnAfficher.Location = New System.Drawing.Point(356, 86)
        Me.btnAfficher.Name = "btnAfficher"
        Me.btnAfficher.Size = New System.Drawing.Size(110, 26)
        Me.btnAfficher.TabIndex = 6
        Me.btnAfficher.Text = "Afficher"
        Me.btnAfficher.UseVisualStyleBackColor = True
        '
        'lblDisponible
        '
        Me.lblDisponible.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblDisponible.ForeColor = System.Drawing.SystemColors.GrayText
        Me.lblDisponible.Location = New System.Drawing.Point(848, 92)
        Me.lblDisponible.Name = "lblDisponible"
        Me.lblDisponible.Size = New System.Drawing.Size(300, 20)
        Me.lblDisponible.TabIndex = 9
        '
        'btnMois
        '
        Me.btnMois.Location = New System.Drawing.Point(472, 86)
        Me.btnMois.Name = "btnMois"
        Me.btnMois.Size = New System.Drawing.Size(120, 26)
        Me.btnMois.TabIndex = 7
        Me.btnMois.Text = "Mois complet"
        Me.btnMois.UseVisualStyleBackColor = True
        '
        'chkASigner
        '
        Me.chkASigner.Location = New System.Drawing.Point(600, 90)
        Me.chkASigner.Name = "chkASigner"
        Me.chkASigner.Size = New System.Drawing.Size(240, 22)
        Me.chkASigner.TabIndex = 8
        Me.chkASigner.Text = "Édition à signer (cartouches)"
        Me.chkASigner.UseVisualStyleBackColor = True
        '
        'lblAvertissement
        '
        Me.lblAvertissement.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblAvertissement.ForeColor = System.Drawing.Color.Firebrick
        Me.lblAvertissement.Location = New System.Drawing.Point(12, 118)
        Me.lblAvertissement.Name = "lblAvertissement"
        Me.lblAvertissement.Size = New System.Drawing.Size(1136, 34)
        Me.lblAvertissement.TabIndex = 10
        '
        'lblSynthese
        '
        Me.lblSynthese.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold)
        Me.lblSynthese.Location = New System.Drawing.Point(12, 156)
        Me.lblSynthese.Name = "lblSynthese"
        Me.lblSynthese.Size = New System.Drawing.Size(620, 20)
        Me.lblSynthese.TabIndex = 11
        Me.lblSynthese.Text = "Répartition sur la période"
        '
        'dgvSynthese
        '
        Me.dgvSynthese.AllowUserToAddRows = False
        Me.dgvSynthese.AllowUserToDeleteRows = False
        Me.dgvSynthese.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvSynthese.Location = New System.Drawing.Point(12, 178)
        Me.dgvSynthese.MultiSelect = False
        Me.dgvSynthese.Name = "dgvSynthese"
        Me.dgvSynthese.ReadOnly = True
        Me.dgvSynthese.RowHeadersWidth = 25
        Me.dgvSynthese.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvSynthese.Size = New System.Drawing.Size(620, 132)
        Me.dgvSynthese.TabIndex = 12
        '
        'lblComptes
        '
        Me.lblComptes.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblComptes.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold)
        Me.lblComptes.Location = New System.Drawing.Point(644, 156)
        Me.lblComptes.Name = "lblComptes"
        Me.lblComptes.Size = New System.Drawing.Size(504, 20)
        Me.lblComptes.TabIndex = 13
        Me.lblComptes.Text = "Comptes bancaires crédités, d'après les pièces conservées"
        '
        'dgvComptes
        '
        Me.dgvComptes.AllowUserToAddRows = False
        Me.dgvComptes.AllowUserToDeleteRows = False
        Me.dgvComptes.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvComptes.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvComptes.Location = New System.Drawing.Point(644, 178)
        Me.dgvComptes.MultiSelect = False
        Me.dgvComptes.Name = "dgvComptes"
        Me.dgvComptes.ReadOnly = True
        Me.dgvComptes.RowHeadersWidth = 25
        Me.dgvComptes.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvComptes.Size = New System.Drawing.Size(504, 132)
        Me.dgvComptes.TabIndex = 14
        '
        'lblParJour
        '
        Me.lblParJour.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold)
        Me.lblParJour.Location = New System.Drawing.Point(12, 320)
        Me.lblParJour.Name = "lblParJour"
        Me.lblParJour.Size = New System.Drawing.Size(1136, 20)
        Me.lblParJour.TabIndex = 15
        Me.lblParJour.Text = "Jour par jour"
        '
        'dgvParJour
        '
        Me.dgvParJour.AllowUserToAddRows = False
        Me.dgvParJour.AllowUserToDeleteRows = False
        Me.dgvParJour.Anchor = CType(((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) Or System.Windows.Forms.AnchorStyles.Right)), System.Windows.Forms.AnchorStyles)
        Me.dgvParJour.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvParJour.Location = New System.Drawing.Point(12, 342)
        Me.dgvParJour.MultiSelect = False
        Me.dgvParJour.Name = "dgvParJour"
        Me.dgvParJour.ReadOnly = True
        Me.dgvParJour.RowHeadersWidth = 25
        Me.dgvParJour.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvParJour.Size = New System.Drawing.Size(1136, 240)
        Me.dgvParJour.TabIndex = 16
        '
        'lblControle
        '
        Me.lblControle.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblControle.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold)
        Me.lblControle.Location = New System.Drawing.Point(12, 590)
        Me.lblControle.Name = "lblControle"
        Me.lblControle.Size = New System.Drawing.Size(1136, 40)
        Me.lblControle.TabIndex = 17
        '
        'btnExporter
        '
        Me.btnExporter.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnExporter.Location = New System.Drawing.Point(12, 640)
        Me.btnExporter.Name = "btnExporter"
        Me.btnExporter.Size = New System.Drawing.Size(200, 32)
        Me.btnExporter.TabIndex = 18
        Me.btnExporter.Text = "Exporter en PDF…"
        Me.btnExporter.UseVisualStyleBackColor = True
        '
        'lblStatut
        '
        Me.lblStatut.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblStatut.ForeColor = System.Drawing.SystemColors.GrayText
        Me.lblStatut.Location = New System.Drawing.Point(222, 646)
        Me.lblStatut.Name = "lblStatut"
        Me.lblStatut.Size = New System.Drawing.Size(822, 20)
        Me.lblStatut.TabIndex = 19
        '
        'btnFermer
        '
        Me.btnFermer.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFermer.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnFermer.Location = New System.Drawing.Point(1056, 640)
        Me.btnFermer.Name = "btnFermer"
        Me.btnFermer.Size = New System.Drawing.Size(92, 32)
        Me.btnFermer.TabIndex = 20
        Me.btnFermer.Text = "Fermer"
        Me.btnFermer.UseVisualStyleBackColor = True
        '
        'sfdEtat
        '
        Me.sfdEtat.Filter = "Document PDF (*.pdf)|*.pdf"
        Me.sfdEtat.Title = "Exporter l'état en PDF"
        '
        'FrmCommissionsBanque
        '
        Me.AcceptButton = Me.btnAfficher
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnFermer
        Me.ClientSize = New System.Drawing.Size(1160, 690)
        Me.Controls.Add(Me.btnFermer)
        Me.Controls.Add(Me.lblStatut)
        Me.Controls.Add(Me.btnExporter)
        Me.Controls.Add(Me.lblControle)
        Me.Controls.Add(Me.dgvParJour)
        Me.Controls.Add(Me.lblParJour)
        Me.Controls.Add(Me.dgvComptes)
        Me.Controls.Add(Me.lblComptes)
        Me.Controls.Add(Me.dgvSynthese)
        Me.Controls.Add(Me.lblSynthese)
        Me.Controls.Add(Me.lblAvertissement)
        Me.Controls.Add(Me.chkASigner)
        Me.Controls.Add(Me.btnMois)
        Me.Controls.Add(Me.lblDisponible)
        Me.Controls.Add(Me.btnAfficher)
        Me.Controls.Add(Me.dtpFin)
        Me.Controls.Add(Me.lblAu)
        Me.Controls.Add(Me.dtpDebut)
        Me.Controls.Add(Me.lblDu)
        Me.Controls.Add(Me.lblSousTitre)
        Me.Controls.Add(Me.lblTitre)
        Me.MinimumSize = New System.Drawing.Size(1040, 620)
        Me.Name = "FrmCommissionsBanque"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Commissions encaissées par la banque"
        CType(Me.dgvSynthese, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvComptes, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvParJour, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents lblTitre As System.Windows.Forms.Label
    Friend WithEvents lblSousTitre As System.Windows.Forms.Label
    Friend WithEvents lblDu As System.Windows.Forms.Label
    Friend WithEvents dtpDebut As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblAu As System.Windows.Forms.Label
    Friend WithEvents dtpFin As System.Windows.Forms.DateTimePicker
    Friend WithEvents btnAfficher As System.Windows.Forms.Button
    Friend WithEvents lblDisponible As System.Windows.Forms.Label
    Friend WithEvents btnMois As System.Windows.Forms.Button
    Friend WithEvents chkASigner As System.Windows.Forms.CheckBox
    Friend WithEvents lblAvertissement As System.Windows.Forms.Label
    Friend WithEvents lblSynthese As System.Windows.Forms.Label
    Friend WithEvents dgvSynthese As System.Windows.Forms.DataGridView
    Friend WithEvents lblComptes As System.Windows.Forms.Label
    Friend WithEvents dgvComptes As System.Windows.Forms.DataGridView
    Friend WithEvents lblParJour As System.Windows.Forms.Label
    Friend WithEvents dgvParJour As System.Windows.Forms.DataGridView
    Friend WithEvents lblControle As System.Windows.Forms.Label
    Friend WithEvents btnExporter As System.Windows.Forms.Button
    Friend WithEvents lblStatut As System.Windows.Forms.Label
    Friend WithEvents btnFermer As System.Windows.Forms.Button
    Friend WithEvents sfdEtat As System.Windows.Forms.SaveFileDialog

End Class
