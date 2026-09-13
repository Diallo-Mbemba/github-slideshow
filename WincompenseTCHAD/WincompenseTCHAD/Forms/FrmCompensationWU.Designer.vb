Option Strict On
Option Explicit On

Partial Class FrmCompensationWU
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.btnActivite = New System.Windows.Forms.Button()
        Me.btnReglement = New System.Windows.Forms.Button()
        Me.btnAfficher = New System.Windows.Forms.Button()
        Me.btnGenererPiece = New System.Windows.Forms.Button()
        Me.btnPieceAccount = New System.Windows.Forms.Button()
        Me.btnParametres = New System.Windows.Forms.Button()
        Me.btnSousAgents = New System.Windows.Forms.Button()
        Me.btnAgences = New System.Windows.Forms.Button()
        Me.btnGroupes = New System.Windows.Forms.Button()
        Me.lblActivite = New System.Windows.Forms.Label()
        Me.lblReglement = New System.Windows.Forms.Label()
        Me.dgvControle = New System.Windows.Forms.DataGridView()
        Me.progressBarTraitement = New System.Windows.Forms.ProgressBar()
        Me.statusStripPrincipal = New System.Windows.Forms.StatusStrip()
        Me.tsslLignesActivite = New System.Windows.Forms.ToolStripStatusLabel()
        Me.tsslLignesReglement = New System.Windows.Forms.ToolStripStatusLabel()
        Me.tsslNombreAccounts = New System.Windows.Forms.ToolStripStatusLabel()
        Me.tsslStatut = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ofdActivite = New System.Windows.Forms.OpenFileDialog()
        Me.ofdReglement = New System.Windows.Forms.OpenFileDialog()
        Me.sfdPieceExcel = New System.Windows.Forms.SaveFileDialog()
        CType(Me.dgvControle, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.statusStripPrincipal.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnActivite
        '
        Me.btnActivite.Location = New System.Drawing.Point(12, 12)
        Me.btnActivite.Name = "btnActivite"
        Me.btnActivite.Size = New System.Drawing.Size(180, 32)
        Me.btnActivite.TabIndex = 0
        Me.btnActivite.Text = "Charger rapport activité"
        Me.btnActivite.UseVisualStyleBackColor = True
        '
        'btnReglement
        '
        Me.btnReglement.Location = New System.Drawing.Point(198, 12)
        Me.btnReglement.Name = "btnReglement"
        Me.btnReglement.Size = New System.Drawing.Size(180, 32)
        Me.btnReglement.TabIndex = 1
        Me.btnReglement.Text = "Charger rapport règlement"
        Me.btnReglement.UseVisualStyleBackColor = True
        '
        'btnAfficher
        '
        Me.btnAfficher.Enabled = False
        Me.btnAfficher.Location = New System.Drawing.Point(600, 12)
        Me.btnAfficher.Name = "btnAfficher"
        Me.btnAfficher.Size = New System.Drawing.Size(160, 32)
        Me.btnAfficher.TabIndex = 2
        Me.btnAfficher.Text = "Afficher / Calculer"
        Me.btnAfficher.UseVisualStyleBackColor = True
        '
        'btnGenererPiece
        '
        Me.btnGenererPiece.Enabled = False
        Me.btnGenererPiece.Location = New System.Drawing.Point(766, 12)
        Me.btnGenererPiece.Name = "btnGenererPiece"
        Me.btnGenererPiece.Size = New System.Drawing.Size(180, 32)
        Me.btnGenererPiece.TabIndex = 3
        Me.btnGenererPiece.Text = "Générer pièce comptable"
        Me.btnGenererPiece.UseVisualStyleBackColor = True
        '
        'btnPieceAccount
        '
        Me.btnPieceAccount.Enabled = False
        Me.btnPieceAccount.Location = New System.Drawing.Point(952, 12)
        Me.btnPieceAccount.Name = "btnPieceAccount"
        Me.btnPieceAccount.Size = New System.Drawing.Size(210, 32)
        Me.btnPieceAccount.TabIndex = 4
        Me.btnPieceAccount.Text = "Pièce de l'Account sélectionné"
        Me.btnPieceAccount.UseVisualStyleBackColor = True
        '
        'btnParametres
        '
        Me.btnParametres.Location = New System.Drawing.Point(566, 50)
        Me.btnParametres.Name = "btnParametres"
        Me.btnParametres.Size = New System.Drawing.Size(143, 26)
        Me.btnParametres.TabIndex = 5
        Me.btnParametres.Text = "Comptes systèmes..."
        Me.btnParametres.UseVisualStyleBackColor = True
        '
        'btnGroupes
        '
        Me.btnGroupes.Location = New System.Drawing.Point(717, 50)
        Me.btnGroupes.Name = "btnGroupes"
        Me.btnGroupes.Size = New System.Drawing.Size(143, 26)
        Me.btnGroupes.TabIndex = 8
        Me.btnGroupes.Text = "Groupes..."
        Me.btnGroupes.UseVisualStyleBackColor = True
        '
        'btnSousAgents
        '
        Me.btnSousAgents.Location = New System.Drawing.Point(868, 50)
        Me.btnSousAgents.Name = "btnSousAgents"
        Me.btnSousAgents.Size = New System.Drawing.Size(143, 26)
        Me.btnSousAgents.TabIndex = 6
        Me.btnSousAgents.Text = "Sous-agents..."
        Me.btnSousAgents.UseVisualStyleBackColor = True
        '
        'btnAgences
        '
        Me.btnAgences.Location = New System.Drawing.Point(1019, 50)
        Me.btnAgences.Name = "btnAgences"
        Me.btnAgences.Size = New System.Drawing.Size(143, 26)
        Me.btnAgences.TabIndex = 7
        Me.btnAgences.Text = "Agences propres..."
        Me.btnAgences.UseVisualStyleBackColor = True
        '
        'lblActivite
        '
        Me.lblActivite.AutoEllipsis = True
        Me.lblActivite.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblActivite.Location = New System.Drawing.Point(12, 50)
        Me.lblActivite.Name = "lblActivite"
        Me.lblActivite.Size = New System.Drawing.Size(270, 23)
        Me.lblActivite.TabIndex = 4
        Me.lblActivite.Text = "(aucun fichier sélectionné)"
        Me.lblActivite.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblReglement
        '
        Me.lblReglement.AutoEllipsis = True
        Me.lblReglement.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblReglement.Location = New System.Drawing.Point(288, 50)
        Me.lblReglement.Name = "lblReglement"
        Me.lblReglement.Size = New System.Drawing.Size(270, 23)
        Me.lblReglement.TabIndex = 5
        Me.lblReglement.Text = "(aucun fichier sélectionné)"
        Me.lblReglement.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'dgvControle
        '
        Me.dgvControle.AllowUserToAddRows = False
        Me.dgvControle.AllowUserToDeleteRows = False
        Me.dgvControle.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvControle.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvControle.Location = New System.Drawing.Point(12, 90)
        Me.dgvControle.Name = "dgvControle"
        Me.dgvControle.ReadOnly = True
        Me.dgvControle.RowHeadersWidth = 25
        Me.dgvControle.Size = New System.Drawing.Size(1160, 480)
        Me.dgvControle.TabIndex = 6
        '
        'progressBarTraitement
        '
        Me.progressBarTraitement.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.progressBarTraitement.Location = New System.Drawing.Point(12, 578)
        Me.progressBarTraitement.Name = "progressBarTraitement"
        Me.progressBarTraitement.Size = New System.Drawing.Size(1160, 18)
        Me.progressBarTraitement.TabIndex = 7
        '
        'statusStripPrincipal
        '
        Me.statusStripPrincipal.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsslLignesActivite, Me.tsslLignesReglement, Me.tsslNombreAccounts, Me.tsslStatut})
        Me.statusStripPrincipal.Location = New System.Drawing.Point(0, 607)
        Me.statusStripPrincipal.Name = "statusStripPrincipal"
        Me.statusStripPrincipal.Size = New System.Drawing.Size(1184, 22)
        Me.statusStripPrincipal.TabIndex = 8
        '
        'tsslLignesActivite
        '
        Me.tsslLignesActivite.Name = "tsslLignesActivite"
        Me.tsslLignesActivite.Size = New System.Drawing.Size(120, 17)
        Me.tsslLignesActivite.Text = "Lignes activité : 0"
        '
        'tsslLignesReglement
        '
        Me.tsslLignesReglement.Name = "tsslLignesReglement"
        Me.tsslLignesReglement.Size = New System.Drawing.Size(130, 17)
        Me.tsslLignesReglement.Text = "Lignes règlement : 0"
        '
        'tsslNombreAccounts
        '
        Me.tsslNombreAccounts.Name = "tsslNombreAccounts"
        Me.tsslNombreAccounts.Size = New System.Drawing.Size(90, 17)
        Me.tsslNombreAccounts.Text = "Accounts : 0"
        '
        'tsslStatut
        '
        Me.tsslStatut.Name = "tsslStatut"
        Me.tsslStatut.Size = New System.Drawing.Size(64, 17)
        Me.tsslStatut.Spring = True
        Me.tsslStatut.Text = "Prêt."
        Me.tsslStatut.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'ofdActivite
        '
        Me.ofdActivite.Filter = "Rapports Western Union (*.zip;*.txt)|*.zip;*.txt|Archives ZIP (*.zip)|*.zip|Fichiers texte (*.txt)|*.txt|Tous les fichiers (*.*)|*.*"
        Me.ofdActivite.Title = "Sélectionner le rapport d'activité Western Union (archive ZIP ou fichier texte)"
        '
        'ofdReglement
        '
        Me.ofdReglement.Filter = "Rapports Western Union (*.zip;*.txt)|*.zip;*.txt|Archives ZIP (*.zip)|*.zip|Fichiers texte (*.txt)|*.txt|Tous les fichiers (*.*)|*.*"
        Me.ofdReglement.Title = "Sélectionner le rapport de règlement Western Union (archive ZIP ou fichier texte)"
        '
        'sfdPieceExcel
        '
        Me.sfdPieceExcel.Filter = "Classeur Excel (*.xlsx)|*.xlsx"
        Me.sfdPieceExcel.Title = "Exporter la pièce comptable"
        '
        'FrmCompensationWU
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1184, 629)
        Me.Controls.Add(Me.statusStripPrincipal)
        Me.Controls.Add(Me.progressBarTraitement)
        Me.Controls.Add(Me.dgvControle)
        Me.Controls.Add(Me.lblReglement)
        Me.Controls.Add(Me.lblActivite)
        Me.Controls.Add(Me.btnPieceAccount)
        Me.Controls.Add(Me.btnParametres)
        Me.Controls.Add(Me.btnSousAgents)
        Me.Controls.Add(Me.btnAgences)
        Me.Controls.Add(Me.btnGroupes)
        Me.Controls.Add(Me.btnGenererPiece)
        Me.Controls.Add(Me.btnAfficher)
        Me.Controls.Add(Me.btnReglement)
        Me.Controls.Add(Me.btnActivite)
        Me.MinimumSize = New System.Drawing.Size(900, 500)
        Me.Name = "FrmCompensationWU"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Compensation Western Union J+1 - Tchad"
        CType(Me.dgvControle, System.ComponentModel.ISupportInitialize).EndInit()
        Me.statusStripPrincipal.ResumeLayout(False)
        Me.statusStripPrincipal.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents btnActivite As System.Windows.Forms.Button
    Friend WithEvents btnReglement As System.Windows.Forms.Button
    Friend WithEvents btnAfficher As System.Windows.Forms.Button
    Friend WithEvents btnGenererPiece As System.Windows.Forms.Button
    Friend WithEvents btnPieceAccount As System.Windows.Forms.Button
    Friend WithEvents btnParametres As System.Windows.Forms.Button
    Friend WithEvents btnSousAgents As System.Windows.Forms.Button
    Friend WithEvents btnAgences As System.Windows.Forms.Button
    Friend WithEvents btnGroupes As System.Windows.Forms.Button
    Friend WithEvents lblActivite As System.Windows.Forms.Label
    Friend WithEvents lblReglement As System.Windows.Forms.Label
    Friend WithEvents dgvControle As System.Windows.Forms.DataGridView
    Friend WithEvents progressBarTraitement As System.Windows.Forms.ProgressBar
    Friend WithEvents statusStripPrincipal As System.Windows.Forms.StatusStrip
    Friend WithEvents tsslLignesActivite As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tsslLignesReglement As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tsslNombreAccounts As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tsslStatut As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents ofdActivite As System.Windows.Forms.OpenFileDialog
    Friend WithEvents ofdReglement As System.Windows.Forms.OpenFileDialog
    Friend WithEvents sfdPieceExcel As System.Windows.Forms.SaveFileDialog

End Class
