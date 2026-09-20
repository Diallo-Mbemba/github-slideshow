Option Strict On
Option Explicit On

Partial Class FrmBordereauJournee
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
        Me.lblJournee = New System.Windows.Forms.Label()
        Me.lblVisa = New System.Windows.Forms.Label()
        Me.lblAlerte = New System.Windows.Forms.Label()
        Me.lblRecapitulatif = New System.Windows.Forms.Label()
        Me.dgvRecapitulatif = New System.Windows.Forms.DataGridView()
        Me.lblEcartes = New System.Windows.Forms.Label()
        Me.dgvEcartes = New System.Windows.Forms.DataGridView()
        Me.btnExporter = New System.Windows.Forms.Button()
        Me.btnViser = New System.Windows.Forms.Button()
        Me.lblStatut = New System.Windows.Forms.Label()
        Me.btnFermer = New System.Windows.Forms.Button()
        Me.sfdBordereau = New System.Windows.Forms.SaveFileDialog()
        CType(Me.dgvRecapitulatif, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvEcartes, System.ComponentModel.ISupportInitialize).BeginInit()
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
        Me.lblTitre.Text = "   Bordereau de compensation"
        Me.lblTitre.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblJournee
        '
        Me.lblJournee.Font = New System.Drawing.Font("Segoe UI", 10.0!, System.Drawing.FontStyle.Bold)
        Me.lblJournee.Location = New System.Drawing.Point(12, 52)
        Me.lblJournee.Name = "lblJournee"
        Me.lblJournee.Size = New System.Drawing.Size(956, 24)
        Me.lblJournee.TabIndex = 1
        Me.lblJournee.Text = "Journée du"
        '
        'lblVisa
        '
        Me.lblVisa.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold)
        Me.lblVisa.Location = New System.Drawing.Point(12, 78)
        Me.lblVisa.Name = "lblVisa"
        Me.lblVisa.Size = New System.Drawing.Size(956, 22)
        Me.lblVisa.TabIndex = 2
        '
        'lblAlerte
        '
        Me.lblAlerte.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblAlerte.ForeColor = System.Drawing.Color.Firebrick
        Me.lblAlerte.Location = New System.Drawing.Point(12, 102)
        Me.lblAlerte.Name = "lblAlerte"
        Me.lblAlerte.Size = New System.Drawing.Size(956, 34)
        Me.lblAlerte.TabIndex = 3
        '
        'lblRecapitulatif
        '
        Me.lblRecapitulatif.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold)
        Me.lblRecapitulatif.Location = New System.Drawing.Point(12, 140)
        Me.lblRecapitulatif.Name = "lblRecapitulatif"
        Me.lblRecapitulatif.Size = New System.Drawing.Size(956, 20)
        Me.lblRecapitulatif.TabIndex = 4
        Me.lblRecapitulatif.Text = "Le traitement de la journée"
        '
        'dgvRecapitulatif
        '
        Me.dgvRecapitulatif.AllowUserToAddRows = False
        Me.dgvRecapitulatif.AllowUserToDeleteRows = False
        Me.dgvRecapitulatif.Anchor = CType(((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) Or System.Windows.Forms.AnchorStyles.Right)), System.Windows.Forms.AnchorStyles)
        Me.dgvRecapitulatif.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvRecapitulatif.Location = New System.Drawing.Point(12, 162)
        Me.dgvRecapitulatif.MultiSelect = False
        Me.dgvRecapitulatif.Name = "dgvRecapitulatif"
        Me.dgvRecapitulatif.ReadOnly = True
        Me.dgvRecapitulatif.RowHeadersWidth = 25
        Me.dgvRecapitulatif.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvRecapitulatif.Size = New System.Drawing.Size(956, 274)
        Me.dgvRecapitulatif.TabIndex = 5
        '
        'lblEcartes
        '
        Me.lblEcartes.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblEcartes.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold)
        Me.lblEcartes.Location = New System.Drawing.Point(12, 444)
        Me.lblEcartes.Name = "lblEcartes"
        Me.lblEcartes.Size = New System.Drawing.Size(956, 20)
        Me.lblEcartes.TabIndex = 6
        Me.lblEcartes.Text = "Accounts non comptabilisés"
        '
        'dgvEcartes
        '
        Me.dgvEcartes.AllowUserToAddRows = False
        Me.dgvEcartes.AllowUserToDeleteRows = False
        Me.dgvEcartes.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvEcartes.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvEcartes.Location = New System.Drawing.Point(12, 466)
        Me.dgvEcartes.MultiSelect = False
        Me.dgvEcartes.Name = "dgvEcartes"
        Me.dgvEcartes.ReadOnly = True
        Me.dgvEcartes.RowHeadersWidth = 25
        Me.dgvEcartes.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvEcartes.Size = New System.Drawing.Size(956, 144)
        Me.dgvEcartes.TabIndex = 7
        '
        'btnExporter
        '
        Me.btnExporter.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnExporter.Location = New System.Drawing.Point(12, 624)
        Me.btnExporter.Name = "btnExporter"
        Me.btnExporter.Size = New System.Drawing.Size(240, 32)
        Me.btnExporter.TabIndex = 8
        Me.btnExporter.Text = "Imprimer le bordereau (PDF)…"
        Me.btnExporter.UseVisualStyleBackColor = True
        '
        'btnViser
        '
        Me.btnViser.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnViser.Location = New System.Drawing.Point(262, 624)
        Me.btnViser.Name = "btnViser"
        Me.btnViser.Size = New System.Drawing.Size(200, 32)
        Me.btnViser.TabIndex = 9
        Me.btnViser.Text = "Viser cette journée…"
        Me.btnViser.UseVisualStyleBackColor = True
        '
        'lblStatut
        '
        Me.lblStatut.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblStatut.ForeColor = System.Drawing.SystemColors.GrayText
        Me.lblStatut.Location = New System.Drawing.Point(472, 630)
        Me.lblStatut.Name = "lblStatut"
        Me.lblStatut.Size = New System.Drawing.Size(392, 20)
        Me.lblStatut.TabIndex = 10
        '
        'btnFermer
        '
        Me.btnFermer.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFermer.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnFermer.Location = New System.Drawing.Point(876, 624)
        Me.btnFermer.Name = "btnFermer"
        Me.btnFermer.Size = New System.Drawing.Size(92, 32)
        Me.btnFermer.TabIndex = 11
        Me.btnFermer.Text = "Fermer"
        Me.btnFermer.UseVisualStyleBackColor = True
        '
        'sfdBordereau
        '
        Me.sfdBordereau.Filter = "Document PDF (*.pdf)|*.pdf"
        Me.sfdBordereau.Title = "Imprimer le bordereau de compensation"
        '
        'FrmBordereauJournee
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnFermer
        Me.ClientSize = New System.Drawing.Size(980, 670)
        Me.Controls.Add(Me.btnFermer)
        Me.Controls.Add(Me.lblStatut)
        Me.Controls.Add(Me.btnViser)
        Me.Controls.Add(Me.btnExporter)
        Me.Controls.Add(Me.dgvEcartes)
        Me.Controls.Add(Me.lblEcartes)
        Me.Controls.Add(Me.dgvRecapitulatif)
        Me.Controls.Add(Me.lblRecapitulatif)
        Me.Controls.Add(Me.lblAlerte)
        Me.Controls.Add(Me.lblVisa)
        Me.Controls.Add(Me.lblJournee)
        Me.Controls.Add(Me.lblTitre)
        Me.MinimumSize = New System.Drawing.Size(860, 600)
        Me.Name = "FrmBordereauJournee"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Bordereau de compensation"
        CType(Me.dgvRecapitulatif, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvEcartes, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents lblTitre As System.Windows.Forms.Label
    Friend WithEvents lblJournee As System.Windows.Forms.Label
    Friend WithEvents lblVisa As System.Windows.Forms.Label
    Friend WithEvents lblAlerte As System.Windows.Forms.Label
    Friend WithEvents lblRecapitulatif As System.Windows.Forms.Label
    Friend WithEvents dgvRecapitulatif As System.Windows.Forms.DataGridView
    Friend WithEvents lblEcartes As System.Windows.Forms.Label
    Friend WithEvents dgvEcartes As System.Windows.Forms.DataGridView
    Friend WithEvents btnExporter As System.Windows.Forms.Button
    Friend WithEvents btnViser As System.Windows.Forms.Button
    Friend WithEvents lblStatut As System.Windows.Forms.Label
    Friend WithEvents btnFermer As System.Windows.Forms.Button
    Friend WithEvents sfdBordereau As System.Windows.Forms.SaveFileDialog

End Class
