Option Strict On
Option Explicit On

Partial Class FrmGroupesStatistiques
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
        Me.lblRecherche = New System.Windows.Forms.Label()
        Me.lblNombre = New System.Windows.Forms.Label()
        Me.lblStatut = New System.Windows.Forms.Label()
        Me.txtRecherche = New System.Windows.Forms.TextBox()
        Me.btnActualiser = New System.Windows.Forms.Button()
        Me.btnNouveau = New System.Windows.Forms.Button()
        Me.btnEnregistrer = New System.Windows.Forms.Button()
        Me.btnSupprimer = New System.Windows.Forms.Button()
        Me.btnSynchroniser = New System.Windows.Forms.Button()
        Me.btnFermer = New System.Windows.Forms.Button()
        Me.dgvListe = New System.Windows.Forms.DataGridView()
        Me.grpDetail = New System.Windows.Forms.GroupBox()
        Me.lblNom = New System.Windows.Forms.Label()
        Me.txtNom = New System.Windows.Forms.TextBox()
        Me.lblCompteActivite = New System.Windows.Forms.Label()
        Me.txtCompteActivite = New System.Windows.Forms.TextBox()
        Me.lblCompteCommission = New System.Windows.Forms.Label()
        Me.txtCompteCommission = New System.Windows.Forms.TextBox()
        Me.lblTaux = New System.Windows.Forms.Label()
        Me.txtTaux = New System.Windows.Forms.TextBox()
        CType(Me.dgvListe, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpDetail.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblRecherche
        '
        Me.lblRecherche.Location = New System.Drawing.Point(12, 16)
        Me.lblRecherche.Name = "lblRecherche"
        Me.lblRecherche.Size = New System.Drawing.Size(80, 20)
        Me.lblRecherche.TabIndex = 0
        Me.lblRecherche.Text = "Rechercher :"
        '
        'txtRecherche
        '
        Me.txtRecherche.Location = New System.Drawing.Point(95, 12)
        Me.txtRecherche.Name = "txtRecherche"
        Me.txtRecherche.Size = New System.Drawing.Size(280, 22)
        Me.txtRecherche.TabIndex = 1
        '
        'btnActualiser
        '
        Me.btnActualiser.Location = New System.Drawing.Point(385, 11)
        Me.btnActualiser.Name = "btnActualiser"
        Me.btnActualiser.Size = New System.Drawing.Size(110, 26)
        Me.btnActualiser.TabIndex = 2
        Me.btnActualiser.Text = "Actualiser"
        Me.btnActualiser.UseVisualStyleBackColor = True
        '
        'lblNombre
        '
        Me.lblNombre.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblNombre.Location = New System.Drawing.Point(508, 16)
        Me.lblNombre.Name = "lblNombre"
        Me.lblNombre.Size = New System.Drawing.Size(300, 20)
        Me.lblNombre.TabIndex = 3
        Me.lblNombre.Text = ""
        Me.lblNombre.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'dgvListe
        '
        Me.dgvListe.AllowUserToAddRows = False
        Me.dgvListe.AllowUserToDeleteRows = False
        Me.dgvListe.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvListe.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvListe.Location = New System.Drawing.Point(12, 45)
        Me.dgvListe.MultiSelect = False
        Me.dgvListe.Name = "dgvListe"
        Me.dgvListe.ReadOnly = True
        Me.dgvListe.RowHeadersWidth = 25
        Me.dgvListe.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvListe.Size = New System.Drawing.Size(796, 250)
        Me.dgvListe.TabIndex = 4
        '
        'grpDetail
        '
        Me.grpDetail.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grpDetail.Controls.Add(Me.lblNom)
        Me.grpDetail.Controls.Add(Me.txtNom)
        Me.grpDetail.Controls.Add(Me.lblCompteActivite)
        Me.grpDetail.Controls.Add(Me.txtCompteActivite)
        Me.grpDetail.Controls.Add(Me.lblCompteCommission)
        Me.grpDetail.Controls.Add(Me.txtCompteCommission)
        Me.grpDetail.Controls.Add(Me.lblTaux)
        Me.grpDetail.Controls.Add(Me.txtTaux)
        Me.grpDetail.Location = New System.Drawing.Point(12, 307)
        Me.grpDetail.Name = "grpDetail"
        Me.grpDetail.Size = New System.Drawing.Size(796, 165)
        Me.grpDetail.TabIndex = 5
        Me.grpDetail.TabStop = False
        Me.grpDetail.Text = "Fiche du groupe statistique"
        '
        'lblNom
        '
        Me.lblNom.Location = New System.Drawing.Point(15, 31)
        Me.lblNom.Name = "lblNom"
        Me.lblNom.Size = New System.Drawing.Size(175, 20)
        Me.lblNom.TabIndex = 6
        Me.lblNom.Text = "Libellé du groupe :"
        '
        'txtNom
        '
        Me.txtNom.Location = New System.Drawing.Point(195, 28)
        Me.txtNom.MaxLength = 255
        Me.txtNom.Name = "txtNom"
        Me.txtNom.Size = New System.Drawing.Size(260, 22)
        Me.txtNom.TabIndex = 7
        '
        'lblCompteActivite
        '
        Me.lblCompteActivite.Location = New System.Drawing.Point(15, 65)
        Me.lblCompteActivite.Name = "lblCompteActivite"
        Me.lblCompteActivite.Size = New System.Drawing.Size(175, 20)
        Me.lblCompteActivite.TabIndex = 8
        Me.lblCompteActivite.Text = "Compte d'activité :"
        '
        'txtCompteActivite
        '
        Me.txtCompteActivite.Location = New System.Drawing.Point(195, 62)
        Me.txtCompteActivite.MaxLength = 255
        Me.txtCompteActivite.Name = "txtCompteActivite"
        Me.txtCompteActivite.Size = New System.Drawing.Size(260, 22)
        Me.txtCompteActivite.TabIndex = 9
        '
        'lblCompteCommission
        '
        Me.lblCompteCommission.Location = New System.Drawing.Point(15, 99)
        Me.lblCompteCommission.Name = "lblCompteCommission"
        Me.lblCompteCommission.Size = New System.Drawing.Size(175, 20)
        Me.lblCompteCommission.TabIndex = 10
        Me.lblCompteCommission.Text = "Compte de commission :"
        '
        'txtCompteCommission
        '
        Me.txtCompteCommission.Location = New System.Drawing.Point(195, 96)
        Me.txtCompteCommission.MaxLength = 255
        Me.txtCompteCommission.Name = "txtCompteCommission"
        Me.txtCompteCommission.Size = New System.Drawing.Size(260, 22)
        Me.txtCompteCommission.TabIndex = 11
        '
        'lblTaux
        '
        Me.lblTaux.Location = New System.Drawing.Point(15, 133)
        Me.lblTaux.Name = "lblTaux"
        Me.lblTaux.Size = New System.Drawing.Size(175, 20)
        Me.lblTaux.TabIndex = 12
        Me.lblTaux.Text = "Taux (0,70 = 70 %) :"
        '
        'txtTaux
        '
        Me.txtTaux.Location = New System.Drawing.Point(195, 130)
        Me.txtTaux.MaxLength = 255
        Me.txtTaux.Name = "txtTaux"
        Me.txtTaux.Size = New System.Drawing.Size(260, 22)
        Me.txtTaux.TabIndex = 13
        '
        'btnNouveau
        '
        Me.btnNouveau.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnNouveau.Location = New System.Drawing.Point(12, 484)
        Me.btnNouveau.Name = "btnNouveau"
        Me.btnNouveau.Size = New System.Drawing.Size(120, 32)
        Me.btnNouveau.TabIndex = 14
        Me.btnNouveau.Text = "Nouveau"
        Me.btnNouveau.UseVisualStyleBackColor = True
        '
        'btnEnregistrer
        '
        Me.btnEnregistrer.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnEnregistrer.Location = New System.Drawing.Point(140, 484)
        Me.btnEnregistrer.Name = "btnEnregistrer"
        Me.btnEnregistrer.Size = New System.Drawing.Size(150, 32)
        Me.btnEnregistrer.TabIndex = 15
        Me.btnEnregistrer.Text = "Enregistrer"
        Me.btnEnregistrer.UseVisualStyleBackColor = True
        '
        'btnSupprimer
        '
        Me.btnSupprimer.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnSupprimer.Location = New System.Drawing.Point(298, 484)
        Me.btnSupprimer.Name = "btnSupprimer"
        Me.btnSupprimer.Size = New System.Drawing.Size(120, 32)
        Me.btnSupprimer.TabIndex = 16
        Me.btnSupprimer.Text = "Supprimer"
        Me.btnSupprimer.UseVisualStyleBackColor = True
        '
        'btnSynchroniser
        '
        Me.btnSynchroniser.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnSynchroniser.Location = New System.Drawing.Point(426, 484)
        Me.btnSynchroniser.Name = "btnSynchroniser"
        Me.btnSynchroniser.Size = New System.Drawing.Size(220, 32)
        Me.btnSynchroniser.TabIndex = 17
        Me.btnSynchroniser.Text = "Synchroniser les sous-agents"
        Me.btnSynchroniser.UseVisualStyleBackColor = True
        '
        'btnFermer
        '
        Me.btnFermer.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFermer.Location = New System.Drawing.Point(688, 484)
        Me.btnFermer.Name = "btnFermer"
        Me.btnFermer.Size = New System.Drawing.Size(120, 32)
        Me.btnFermer.TabIndex = 18
        Me.btnFermer.Text = "Fermer"
        Me.btnFermer.UseVisualStyleBackColor = True
        '
        'lblStatut
        '
        Me.lblStatut.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblStatut.ForeColor = System.Drawing.SystemColors.GrayText
        Me.lblStatut.Location = New System.Drawing.Point(12, 524)
        Me.lblStatut.Name = "lblStatut"
        Me.lblStatut.Size = New System.Drawing.Size(796, 20)
        Me.lblStatut.TabIndex = 19
        Me.lblStatut.Text = ""
        '
        'FrmGroupesStatistiques
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnFermer
        Me.ClientSize = New System.Drawing.Size(820, 554)
        Me.Controls.Add(Me.lblRecherche)
        Me.Controls.Add(Me.txtRecherche)
        Me.Controls.Add(Me.btnActualiser)
        Me.Controls.Add(Me.lblNombre)
        Me.Controls.Add(Me.dgvListe)
        Me.Controls.Add(Me.grpDetail)
        Me.Controls.Add(Me.btnNouveau)
        Me.Controls.Add(Me.btnEnregistrer)
        Me.Controls.Add(Me.btnSupprimer)
        Me.Controls.Add(Me.btnSynchroniser)
        Me.Controls.Add(Me.btnFermer)
        Me.Controls.Add(Me.lblStatut)
        Me.MinimumSize = New System.Drawing.Size(836, 594)
        Me.Name = "FrmGroupesStatistiques"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Groupes statistiques (T_GroupeStatistique)"
        CType(Me.dgvListe, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpDetail.ResumeLayout(False)
        Me.grpDetail.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()
    End Sub

    Friend WithEvents lblRecherche As System.Windows.Forms.Label
    Friend WithEvents lblNombre As System.Windows.Forms.Label
    Friend WithEvents lblStatut As System.Windows.Forms.Label
    Friend WithEvents txtRecherche As System.Windows.Forms.TextBox
    Friend WithEvents btnActualiser As System.Windows.Forms.Button
    Friend WithEvents btnNouveau As System.Windows.Forms.Button
    Friend WithEvents btnEnregistrer As System.Windows.Forms.Button
    Friend WithEvents btnSupprimer As System.Windows.Forms.Button
    Friend WithEvents btnSynchroniser As System.Windows.Forms.Button
    Friend WithEvents btnFermer As System.Windows.Forms.Button
    Friend WithEvents dgvListe As System.Windows.Forms.DataGridView
    Friend WithEvents grpDetail As System.Windows.Forms.GroupBox
    Friend WithEvents lblNom As System.Windows.Forms.Label
    Friend WithEvents txtNom As System.Windows.Forms.TextBox
    Friend WithEvents lblCompteActivite As System.Windows.Forms.Label
    Friend WithEvents txtCompteActivite As System.Windows.Forms.TextBox
    Friend WithEvents lblCompteCommission As System.Windows.Forms.Label
    Friend WithEvents txtCompteCommission As System.Windows.Forms.TextBox
    Friend WithEvents lblTaux As System.Windows.Forms.Label
    Friend WithEvents txtTaux As System.Windows.Forms.TextBox

End Class
