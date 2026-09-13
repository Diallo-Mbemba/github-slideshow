Option Strict On
Option Explicit On

Partial Class FrmSousAgents
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
        Me.btnFermer = New System.Windows.Forms.Button()
        Me.dgvListe = New System.Windows.Forms.DataGridView()
        Me.grpDetail = New System.Windows.Forms.GroupBox()
        Me.lblCodePdv = New System.Windows.Forms.Label()
        Me.txtCodePdv = New System.Windows.Forms.TextBox()
        Me.lblDesignation = New System.Windows.Forms.Label()
        Me.txtDesignation = New System.Windows.Forms.TextBox()
        Me.lblGroupeStatistique = New System.Windows.Forms.Label()
        Me.txtGroupeStatistique = New System.Windows.Forms.TextBox()
        Me.lblTaux = New System.Windows.Forms.Label()
        Me.txtTaux = New System.Windows.Forms.TextBox()
        Me.lblCompteCompense = New System.Windows.Forms.Label()
        Me.txtCompteCompense = New System.Windows.Forms.TextBox()
        Me.lblCompteCommission = New System.Windows.Forms.Label()
        Me.txtCompteCommission = New System.Windows.Forms.TextBox()
        Me.lblCodeAgence = New System.Windows.Forms.Label()
        Me.txtCodeAgence = New System.Windows.Forms.TextBox()
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
        Me.lblNombre.Location = New System.Drawing.Point(588, 16)
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
        Me.dgvListe.Size = New System.Drawing.Size(876, 270)
        Me.dgvListe.TabIndex = 4
        '
        'grpDetail
        '
        Me.grpDetail.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grpDetail.Controls.Add(Me.lblCodePdv)
        Me.grpDetail.Controls.Add(Me.txtCodePdv)
        Me.grpDetail.Controls.Add(Me.lblDesignation)
        Me.grpDetail.Controls.Add(Me.txtDesignation)
        Me.grpDetail.Controls.Add(Me.lblGroupeStatistique)
        Me.grpDetail.Controls.Add(Me.txtGroupeStatistique)
        Me.grpDetail.Controls.Add(Me.lblTaux)
        Me.grpDetail.Controls.Add(Me.txtTaux)
        Me.grpDetail.Controls.Add(Me.lblCompteCompense)
        Me.grpDetail.Controls.Add(Me.txtCompteCompense)
        Me.grpDetail.Controls.Add(Me.lblCompteCommission)
        Me.grpDetail.Controls.Add(Me.txtCompteCommission)
        Me.grpDetail.Controls.Add(Me.lblCodeAgence)
        Me.grpDetail.Controls.Add(Me.txtCodeAgence)
        Me.grpDetail.Location = New System.Drawing.Point(12, 327)
        Me.grpDetail.Name = "grpDetail"
        Me.grpDetail.Size = New System.Drawing.Size(876, 160)
        Me.grpDetail.TabIndex = 5
        Me.grpDetail.TabStop = False
        Me.grpDetail.Text = "Fiche du sous-agent"
        '
        'lblCodePdv
        '
        Me.lblCodePdv.Location = New System.Drawing.Point(15, 31)
        Me.lblCodePdv.Name = "lblCodePdv"
        Me.lblCodePdv.Size = New System.Drawing.Size(155, 20)
        Me.lblCodePdv.TabIndex = 6
        Me.lblCodePdv.Text = "Account (Code_Pdv) :"
        '
        'txtCodePdv
        '
        Me.txtCodePdv.Location = New System.Drawing.Point(175, 28)
        Me.txtCodePdv.MaxLength = 255
        Me.txtCodePdv.Name = "txtCodePdv"
        Me.txtCodePdv.Size = New System.Drawing.Size(230, 22)
        Me.txtCodePdv.TabIndex = 7
        '
        'lblDesignation
        '
        Me.lblDesignation.Location = New System.Drawing.Point(15, 65)
        Me.lblDesignation.Name = "lblDesignation"
        Me.lblDesignation.Size = New System.Drawing.Size(155, 20)
        Me.lblDesignation.TabIndex = 8
        Me.lblDesignation.Text = "Désignation agence :"
        '
        'txtDesignation
        '
        Me.txtDesignation.Location = New System.Drawing.Point(175, 62)
        Me.txtDesignation.MaxLength = 255
        Me.txtDesignation.Name = "txtDesignation"
        Me.txtDesignation.Size = New System.Drawing.Size(230, 22)
        Me.txtDesignation.TabIndex = 9
        '
        'lblGroupeStatistique
        '
        Me.lblGroupeStatistique.Location = New System.Drawing.Point(15, 99)
        Me.lblGroupeStatistique.Name = "lblGroupeStatistique"
        Me.lblGroupeStatistique.Size = New System.Drawing.Size(155, 20)
        Me.lblGroupeStatistique.TabIndex = 10
        Me.lblGroupeStatistique.Text = "Groupe statistique :"
        '
        'txtGroupeStatistique
        '
        Me.txtGroupeStatistique.Location = New System.Drawing.Point(175, 96)
        Me.txtGroupeStatistique.MaxLength = 255
        Me.txtGroupeStatistique.Name = "txtGroupeStatistique"
        Me.txtGroupeStatistique.Size = New System.Drawing.Size(230, 22)
        Me.txtGroupeStatistique.TabIndex = 11
        '
        'lblTaux
        '
        Me.lblTaux.Location = New System.Drawing.Point(15, 133)
        Me.lblTaux.Name = "lblTaux"
        Me.lblTaux.Size = New System.Drawing.Size(155, 20)
        Me.lblTaux.TabIndex = 12
        Me.lblTaux.Text = "Taux (0,70 = 70 %) :"
        '
        'txtTaux
        '
        Me.txtTaux.Location = New System.Drawing.Point(175, 130)
        Me.txtTaux.MaxLength = 255
        Me.txtTaux.Name = "txtTaux"
        Me.txtTaux.Size = New System.Drawing.Size(230, 22)
        Me.txtTaux.TabIndex = 13
        '
        'lblCompteCompense
        '
        Me.lblCompteCompense.Location = New System.Drawing.Point(430, 31)
        Me.lblCompteCompense.Name = "lblCompteCompense"
        Me.lblCompteCompense.Size = New System.Drawing.Size(165, 20)
        Me.lblCompteCompense.TabIndex = 14
        Me.lblCompteCompense.Text = "Compte de compensation :"
        '
        'txtCompteCompense
        '
        Me.txtCompteCompense.Location = New System.Drawing.Point(600, 28)
        Me.txtCompteCompense.MaxLength = 255
        Me.txtCompteCompense.Name = "txtCompteCompense"
        Me.txtCompteCompense.Size = New System.Drawing.Size(255, 22)
        Me.txtCompteCompense.TabIndex = 15
        '
        'lblCompteCommission
        '
        Me.lblCompteCommission.Location = New System.Drawing.Point(430, 65)
        Me.lblCompteCommission.Name = "lblCompteCommission"
        Me.lblCompteCommission.Size = New System.Drawing.Size(165, 20)
        Me.lblCompteCommission.TabIndex = 16
        Me.lblCompteCommission.Text = "Compte de commission :"
        '
        'txtCompteCommission
        '
        Me.txtCompteCommission.Location = New System.Drawing.Point(600, 62)
        Me.txtCompteCommission.MaxLength = 255
        Me.txtCompteCommission.Name = "txtCompteCommission"
        Me.txtCompteCommission.Size = New System.Drawing.Size(255, 22)
        Me.txtCompteCommission.TabIndex = 17
        '
        'lblCodeAgence
        '
        Me.lblCodeAgence.Location = New System.Drawing.Point(430, 99)
        Me.lblCodeAgence.Name = "lblCodeAgence"
        Me.lblCodeAgence.Size = New System.Drawing.Size(165, 20)
        Me.lblCodeAgence.TabIndex = 18
        Me.lblCodeAgence.Text = "Code agence :"
        '
        'txtCodeAgence
        '
        Me.txtCodeAgence.Location = New System.Drawing.Point(600, 96)
        Me.txtCodeAgence.MaxLength = 255
        Me.txtCodeAgence.Name = "txtCodeAgence"
        Me.txtCodeAgence.Size = New System.Drawing.Size(255, 22)
        Me.txtCodeAgence.TabIndex = 19
        '
        'btnNouveau
        '
        Me.btnNouveau.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnNouveau.Location = New System.Drawing.Point(12, 499)
        Me.btnNouveau.Name = "btnNouveau"
        Me.btnNouveau.Size = New System.Drawing.Size(120, 32)
        Me.btnNouveau.TabIndex = 20
        Me.btnNouveau.Text = "Nouveau"
        Me.btnNouveau.UseVisualStyleBackColor = True
        '
        'btnEnregistrer
        '
        Me.btnEnregistrer.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnEnregistrer.Location = New System.Drawing.Point(140, 499)
        Me.btnEnregistrer.Name = "btnEnregistrer"
        Me.btnEnregistrer.Size = New System.Drawing.Size(150, 32)
        Me.btnEnregistrer.TabIndex = 21
        Me.btnEnregistrer.Text = "Enregistrer"
        Me.btnEnregistrer.UseVisualStyleBackColor = True
        '
        'btnSupprimer
        '
        Me.btnSupprimer.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnSupprimer.Location = New System.Drawing.Point(298, 499)
        Me.btnSupprimer.Name = "btnSupprimer"
        Me.btnSupprimer.Size = New System.Drawing.Size(120, 32)
        Me.btnSupprimer.TabIndex = 22
        Me.btnSupprimer.Text = "Supprimer"
        Me.btnSupprimer.UseVisualStyleBackColor = True
        '
        'btnFermer
        '
        Me.btnFermer.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFermer.Location = New System.Drawing.Point(768, 499)
        Me.btnFermer.Name = "btnFermer"
        Me.btnFermer.Size = New System.Drawing.Size(120, 32)
        Me.btnFermer.TabIndex = 23
        Me.btnFermer.Text = "Fermer"
        Me.btnFermer.UseVisualStyleBackColor = True
        '
        'lblStatut
        '
        Me.lblStatut.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblStatut.ForeColor = System.Drawing.SystemColors.GrayText
        Me.lblStatut.Location = New System.Drawing.Point(12, 539)
        Me.lblStatut.Name = "lblStatut"
        Me.lblStatut.Size = New System.Drawing.Size(876, 20)
        Me.lblStatut.TabIndex = 24
        Me.lblStatut.Text = ""
        '
        'FrmSousAgents
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnFermer
        Me.ClientSize = New System.Drawing.Size(900, 569)
        Me.Controls.Add(Me.lblRecherche)
        Me.Controls.Add(Me.txtRecherche)
        Me.Controls.Add(Me.btnActualiser)
        Me.Controls.Add(Me.lblNombre)
        Me.Controls.Add(Me.dgvListe)
        Me.Controls.Add(Me.grpDetail)
        Me.Controls.Add(Me.btnNouveau)
        Me.Controls.Add(Me.btnEnregistrer)
        Me.Controls.Add(Me.btnSupprimer)
        Me.Controls.Add(Me.btnFermer)
        Me.Controls.Add(Me.lblStatut)
        Me.MinimumSize = New System.Drawing.Size(916, 609)
        Me.Name = "FrmSousAgents"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Sous-agents Western Union (T_Pdv_SA)"
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
    Friend WithEvents btnFermer As System.Windows.Forms.Button
    Friend WithEvents dgvListe As System.Windows.Forms.DataGridView
    Friend WithEvents grpDetail As System.Windows.Forms.GroupBox
    Friend WithEvents lblCodePdv As System.Windows.Forms.Label
    Friend WithEvents txtCodePdv As System.Windows.Forms.TextBox
    Friend WithEvents lblDesignation As System.Windows.Forms.Label
    Friend WithEvents txtDesignation As System.Windows.Forms.TextBox
    Friend WithEvents lblGroupeStatistique As System.Windows.Forms.Label
    Friend WithEvents txtGroupeStatistique As System.Windows.Forms.TextBox
    Friend WithEvents lblTaux As System.Windows.Forms.Label
    Friend WithEvents txtTaux As System.Windows.Forms.TextBox
    Friend WithEvents lblCompteCompense As System.Windows.Forms.Label
    Friend WithEvents txtCompteCompense As System.Windows.Forms.TextBox
    Friend WithEvents lblCompteCommission As System.Windows.Forms.Label
    Friend WithEvents txtCompteCommission As System.Windows.Forms.TextBox
    Friend WithEvents lblCodeAgence As System.Windows.Forms.Label
    Friend WithEvents txtCodeAgence As System.Windows.Forms.TextBox

End Class
