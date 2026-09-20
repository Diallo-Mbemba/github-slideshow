Option Strict On
Option Explicit On

Partial Class FrmPiecesArchivees
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
        Me.lblJournees = New System.Windows.Forms.Label()
        Me.dgvJournees = New System.Windows.Forms.DataGridView()
        Me.lblPiece = New System.Windows.Forms.Label()
        Me.dgvPiece = New System.Windows.Forms.DataGridView()
        Me.lblTotaux = New System.Windows.Forms.Label()
        Me.btnActualiser = New System.Windows.Forms.Button()
        Me.btnPiece = New System.Windows.Forms.Button()
        Me.btnCoreBanking = New System.Windows.Forms.Button()
        Me.btnFermer = New System.Windows.Forms.Button()
        Me.btnAnnuler = New System.Windows.Forms.Button()
        Me.btnBordereau = New System.Windows.Forms.Button()
        Me.lblStatut = New System.Windows.Forms.Label()
        CType(Me.dgvJournees, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvPiece, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'lblTitre
        '
        Me.lblTitre.BackColor = System.Drawing.Color.Black
        Me.lblTitre.Font = New System.Drawing.Font("Segoe UI", 14.0!, System.Drawing.FontStyle.Bold)
        Me.lblTitre.ForeColor = System.Drawing.Color.Yellow
        Me.lblTitre.Location = New System.Drawing.Point(12, 12)
        Me.lblTitre.Name = "lblTitre"
        Me.lblTitre.Size = New System.Drawing.Size(1316, 34)
        Me.lblTitre.TabIndex = 0
        Me.lblTitre.Text = "   Pièces comptables conservées"
        Me.lblTitre.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblSousTitre
        '
        Me.lblSousTitre.ForeColor = System.Drawing.SystemColors.GrayText
        Me.lblSousTitre.Location = New System.Drawing.Point(12, 50)
        Me.lblSousTitre.Name = "lblSousTitre"
        Me.lblSousTitre.Size = New System.Drawing.Size(1316, 30)
        Me.lblSousTitre.TabIndex = 1
        Me.lblSousTitre.Text = "Chaque pièce est celle qui a été produite ce jour-là, conservée ligne à ligne. Rien n'est recalculé."
        '
        'lblJournees
        '
        Me.lblJournees.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold)
        Me.lblJournees.Location = New System.Drawing.Point(12, 86)
        Me.lblJournees.Name = "lblJournees"
        Me.lblJournees.Size = New System.Drawing.Size(460, 20)
        Me.lblJournees.TabIndex = 2
        Me.lblJournees.Text = "Journées conservées"
        '
        'dgvJournees
        '
        Me.dgvJournees.AllowUserToAddRows = False
        Me.dgvJournees.AllowUserToDeleteRows = False
        Me.dgvJournees.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left)), System.Windows.Forms.AnchorStyles)
        Me.dgvJournees.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvJournees.Location = New System.Drawing.Point(12, 108)
        Me.dgvJournees.MultiSelect = False
        Me.dgvJournees.Name = "dgvJournees"
        Me.dgvJournees.ReadOnly = True
        Me.dgvJournees.RowHeadersWidth = 25
        Me.dgvJournees.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvJournees.Size = New System.Drawing.Size(460, 500)
        Me.dgvJournees.TabIndex = 3
        '
        'lblPiece
        '
        Me.lblPiece.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblPiece.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold)
        Me.lblPiece.Location = New System.Drawing.Point(484, 86)
        Me.lblPiece.Name = "lblPiece"
        Me.lblPiece.Size = New System.Drawing.Size(844, 20)
        Me.lblPiece.TabIndex = 4
        Me.lblPiece.Text = "Écritures de la journée"
        '
        'dgvPiece
        '
        Me.dgvPiece.AllowUserToAddRows = False
        Me.dgvPiece.AllowUserToDeleteRows = False
        Me.dgvPiece.Anchor = CType(((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) Or System.Windows.Forms.AnchorStyles.Right)), System.Windows.Forms.AnchorStyles)
        Me.dgvPiece.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvPiece.Location = New System.Drawing.Point(484, 108)
        Me.dgvPiece.MultiSelect = False
        Me.dgvPiece.Name = "dgvPiece"
        Me.dgvPiece.ReadOnly = True
        Me.dgvPiece.RowHeadersWidth = 25
        Me.dgvPiece.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvPiece.Size = New System.Drawing.Size(844, 470)
        Me.dgvPiece.TabIndex = 5
        '
        'lblTotaux
        '
        Me.lblTotaux.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblTotaux.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold)
        Me.lblTotaux.Location = New System.Drawing.Point(484, 584)
        Me.lblTotaux.Name = "lblTotaux"
        Me.lblTotaux.Size = New System.Drawing.Size(844, 24)
        Me.lblTotaux.TabIndex = 6
        '
        'btnActualiser
        '
        Me.btnActualiser.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnActualiser.Location = New System.Drawing.Point(12, 620)
        Me.btnActualiser.Name = "btnActualiser"
        Me.btnActualiser.Size = New System.Drawing.Size(120, 32)
        Me.btnActualiser.TabIndex = 7
        Me.btnActualiser.Text = "Actualiser"
        Me.btnActualiser.UseVisualStyleBackColor = True
        '
        'btnPiece
        '
        Me.btnPiece.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnPiece.Location = New System.Drawing.Point(860, 620)
        Me.btnPiece.Name = "btnPiece"
        Me.btnPiece.Size = New System.Drawing.Size(180, 32)
        Me.btnPiece.TabIndex = 10
        Me.btnPiece.Text = "Ouvrir la pièce…"
        Me.btnPiece.UseVisualStyleBackColor = True
        '
        'btnCoreBanking
        '
        Me.btnCoreBanking.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCoreBanking.Location = New System.Drawing.Point(1048, 620)
        Me.btnCoreBanking.Name = "btnCoreBanking"
        Me.btnCoreBanking.Size = New System.Drawing.Size(180, 32)
        Me.btnCoreBanking.TabIndex = 11
        Me.btnCoreBanking.Text = "Fichier core banking…"
        Me.btnCoreBanking.UseVisualStyleBackColor = True
        '
        'btnFermer
        '
        Me.btnFermer.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFermer.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnFermer.Location = New System.Drawing.Point(1236, 620)
        Me.btnFermer.Name = "btnFermer"
        Me.btnFermer.Size = New System.Drawing.Size(92, 32)
        Me.btnFermer.TabIndex = 12
        Me.btnFermer.Text = "Fermer"
        Me.btnFermer.UseVisualStyleBackColor = True
        '
        'btnAnnuler
        '
        Me.btnAnnuler.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnAnnuler.Location = New System.Drawing.Point(140, 620)
        Me.btnAnnuler.Name = "btnAnnuler"
        Me.btnAnnuler.Size = New System.Drawing.Size(256, 32)
        Me.btnAnnuler.TabIndex = 8
        Me.btnAnnuler.Text = "Annuler cette comptabilisation…"
        Me.btnAnnuler.UseVisualStyleBackColor = True
        '
        'btnBordereau
        '
        Me.btnBordereau.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnBordereau.Location = New System.Drawing.Point(404, 620)
        Me.btnBordereau.Name = "btnBordereau"
        Me.btnBordereau.Size = New System.Drawing.Size(200, 32)
        Me.btnBordereau.TabIndex = 9
        Me.btnBordereau.Text = "Bordereau de la journée…"
        Me.btnBordereau.UseVisualStyleBackColor = True
        '
        'lblStatut
        '
        Me.lblStatut.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblStatut.ForeColor = System.Drawing.SystemColors.GrayText
        Me.lblStatut.Location = New System.Drawing.Point(612, 628)
        Me.lblStatut.Name = "lblStatut"
        Me.lblStatut.Size = New System.Drawing.Size(240, 20)
        Me.lblStatut.TabIndex = 13
        '
        'FrmPiecesArchivees
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnFermer
        Me.ClientSize = New System.Drawing.Size(1340, 664)
        Me.Controls.Add(Me.lblStatut)
        Me.Controls.Add(Me.btnFermer)
        Me.Controls.Add(Me.btnCoreBanking)
        Me.Controls.Add(Me.btnPiece)
        Me.Controls.Add(Me.btnBordereau)
        Me.Controls.Add(Me.btnAnnuler)
        Me.Controls.Add(Me.btnActualiser)
        Me.Controls.Add(Me.lblTotaux)
        Me.Controls.Add(Me.dgvPiece)
        Me.Controls.Add(Me.lblPiece)
        Me.Controls.Add(Me.dgvJournees)
        Me.Controls.Add(Me.lblJournees)
        Me.Controls.Add(Me.lblSousTitre)
        Me.Controls.Add(Me.lblTitre)
        Me.MinimumSize = New System.Drawing.Size(1060, 560)
        Me.Name = "FrmPiecesArchivees"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Pièces comptables conservées"
        CType(Me.dgvJournees, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvPiece, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents lblTitre As System.Windows.Forms.Label
    Friend WithEvents lblSousTitre As System.Windows.Forms.Label
    Friend WithEvents lblJournees As System.Windows.Forms.Label
    Friend WithEvents dgvJournees As System.Windows.Forms.DataGridView
    Friend WithEvents lblPiece As System.Windows.Forms.Label
    Friend WithEvents dgvPiece As System.Windows.Forms.DataGridView
    Friend WithEvents lblTotaux As System.Windows.Forms.Label
    Friend WithEvents btnActualiser As System.Windows.Forms.Button
    Friend WithEvents btnAnnuler As System.Windows.Forms.Button
    Friend WithEvents btnBordereau As System.Windows.Forms.Button
    Friend WithEvents btnPiece As System.Windows.Forms.Button
    Friend WithEvents btnCoreBanking As System.Windows.Forms.Button
    Friend WithEvents btnFermer As System.Windows.Forms.Button
    Friend WithEvents lblStatut As System.Windows.Forms.Label

End Class
