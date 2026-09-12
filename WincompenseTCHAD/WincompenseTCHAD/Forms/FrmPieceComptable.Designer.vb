Option Strict On
Option Explicit On

Partial Class FrmPieceComptable
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
        Me.dgvPiece = New System.Windows.Forms.DataGridView()
        Me.lblTotaux = New System.Windows.Forms.Label()
        Me.lblEcart = New System.Windows.Forms.Label()
        Me.btnOuvrirExcel = New System.Windows.Forms.Button()
        Me.btnFermer = New System.Windows.Forms.Button()
        CType(Me.dgvPiece, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'lblTitre
        '
        Me.lblTitre.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblTitre.AutoEllipsis = True
        Me.lblTitre.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.0!, System.Drawing.FontStyle.Bold)
        Me.lblTitre.Location = New System.Drawing.Point(12, 12)
        Me.lblTitre.Name = "lblTitre"
        Me.lblTitre.Size = New System.Drawing.Size(876, 22)
        Me.lblTitre.TabIndex = 0
        Me.lblTitre.Text = "Pièce comptable"
        '
        'lblSousTitre
        '
        Me.lblSousTitre.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblSousTitre.AutoEllipsis = True
        Me.lblSousTitre.ForeColor = System.Drawing.SystemColors.ControlDarkDark
        Me.lblSousTitre.Location = New System.Drawing.Point(12, 38)
        Me.lblSousTitre.Name = "lblSousTitre"
        Me.lblSousTitre.Size = New System.Drawing.Size(876, 18)
        Me.lblSousTitre.TabIndex = 1
        '
        'dgvPiece
        '
        Me.dgvPiece.AllowUserToAddRows = False
        Me.dgvPiece.AllowUserToDeleteRows = False
        Me.dgvPiece.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvPiece.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvPiece.Location = New System.Drawing.Point(12, 64)
        Me.dgvPiece.Name = "dgvPiece"
        Me.dgvPiece.ReadOnly = True
        Me.dgvPiece.RowHeadersWidth = 25
        Me.dgvPiece.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvPiece.Size = New System.Drawing.Size(876, 404)
        Me.dgvPiece.TabIndex = 2
        '
        'lblTotaux
        '
        Me.lblTotaux.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblTotaux.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold)
        Me.lblTotaux.Location = New System.Drawing.Point(12, 480)
        Me.lblTotaux.Name = "lblTotaux"
        Me.lblTotaux.Size = New System.Drawing.Size(600, 18)
        Me.lblTotaux.TabIndex = 3
        '
        'lblEcart
        '
        Me.lblEcart.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblEcart.Location = New System.Drawing.Point(12, 502)
        Me.lblEcart.Name = "lblEcart"
        Me.lblEcart.Size = New System.Drawing.Size(600, 32)
        Me.lblEcart.TabIndex = 4
        '
        'btnOuvrirExcel
        '
        Me.btnOuvrirExcel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOuvrirExcel.Location = New System.Drawing.Point(628, 486)
        Me.btnOuvrirExcel.Name = "btnOuvrirExcel"
        Me.btnOuvrirExcel.Size = New System.Drawing.Size(150, 32)
        Me.btnOuvrirExcel.TabIndex = 5
        Me.btnOuvrirExcel.Text = "Ouvrir dans Excel"
        Me.btnOuvrirExcel.UseVisualStyleBackColor = True
        '
        'btnFermer
        '
        Me.btnFermer.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFermer.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnFermer.Location = New System.Drawing.Point(788, 486)
        Me.btnFermer.Name = "btnFermer"
        Me.btnFermer.Size = New System.Drawing.Size(100, 32)
        Me.btnFermer.TabIndex = 6
        Me.btnFermer.Text = "Fermer"
        Me.btnFermer.UseVisualStyleBackColor = True
        '
        'FrmPieceComptable
        '
        Me.AcceptButton = Me.btnFermer
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnFermer
        Me.ClientSize = New System.Drawing.Size(900, 534)
        Me.Controls.Add(Me.btnFermer)
        Me.Controls.Add(Me.btnOuvrirExcel)
        Me.Controls.Add(Me.lblEcart)
        Me.Controls.Add(Me.lblTotaux)
        Me.Controls.Add(Me.dgvPiece)
        Me.Controls.Add(Me.lblSousTitre)
        Me.Controls.Add(Me.lblTitre)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(700, 400)
        Me.Name = "FrmPieceComptable"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Pièce comptable"
        CType(Me.dgvPiece, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents lblTitre As System.Windows.Forms.Label
    Friend WithEvents lblSousTitre As System.Windows.Forms.Label
    Friend WithEvents dgvPiece As System.Windows.Forms.DataGridView
    Friend WithEvents lblTotaux As System.Windows.Forms.Label
    Friend WithEvents lblEcart As System.Windows.Forms.Label
    Friend WithEvents btnOuvrirExcel As System.Windows.Forms.Button
    Friend WithEvents btnFermer As System.Windows.Forms.Button

End Class
