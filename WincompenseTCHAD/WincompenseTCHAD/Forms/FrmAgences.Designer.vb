Option Strict On
Option Explicit On

Partial Class FrmAgences
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
        Me.lblCodeSite = New System.Windows.Forms.Label()
        Me.txtCodeSite = New System.Windows.Forms.TextBox()
        Me.lblDesignation = New System.Windows.Forms.Label()
        Me.txtDesignation = New System.Windows.Forms.TextBox()
        Me.lblCodeAgenceVoyager = New System.Windows.Forms.Label()
        Me.txtCodeAgenceVoyager = New System.Windows.Forms.TextBox()
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
        Me.lblNombre.Location = New System.Drawing.Point(448, 16)
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
        Me.dgvListe.Size = New System.Drawing.Size(736, 250)
        Me.dgvListe.TabIndex = 4
        '
        'grpDetail
        '
        Me.grpDetail.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grpDetail.Controls.Add(Me.lblCodeSite)
        Me.grpDetail.Controls.Add(Me.txtCodeSite)
        Me.grpDetail.Controls.Add(Me.lblDesignation)
        Me.grpDetail.Controls.Add(Me.txtDesignation)
        Me.grpDetail.Controls.Add(Me.lblCodeAgenceVoyager)
        Me.grpDetail.Controls.Add(Me.txtCodeAgenceVoyager)
        Me.grpDetail.Location = New System.Drawing.Point(12, 307)
        Me.grpDetail.Name = "grpDetail"
        Me.grpDetail.Size = New System.Drawing.Size(736, 130)
        Me.grpDetail.TabIndex = 5
        Me.grpDetail.TabStop = False
        Me.grpDetail.Text = "Fiche de l'agence"
        '
        'lblCodeSite
        '
        Me.lblCodeSite.Location = New System.Drawing.Point(15, 31)
        Me.lblCodeSite.Name = "lblCodeSite"
        Me.lblCodeSite.Size = New System.Drawing.Size(165, 20)
        Me.lblCodeSite.TabIndex = 6
        Me.lblCodeSite.Text = "Account (Codesite) :"
        '
        'txtCodeSite
        '
        Me.txtCodeSite.Location = New System.Drawing.Point(185, 28)
        Me.txtCodeSite.MaxLength = 255
        Me.txtCodeSite.Name = "txtCodeSite"
        Me.txtCodeSite.Size = New System.Drawing.Size(280, 22)
        Me.txtCodeSite.TabIndex = 7
        '
        'lblDesignation
        '
        Me.lblDesignation.Location = New System.Drawing.Point(15, 65)
        Me.lblDesignation.Name = "lblDesignation"
        Me.lblDesignation.Size = New System.Drawing.Size(165, 20)
        Me.lblDesignation.TabIndex = 8
        Me.lblDesignation.Text = "Désignation agence :"
        '
        'txtDesignation
        '
        Me.txtDesignation.Location = New System.Drawing.Point(185, 62)
        Me.txtDesignation.MaxLength = 255
        Me.txtDesignation.Name = "txtDesignation"
        Me.txtDesignation.Size = New System.Drawing.Size(280, 22)
        Me.txtDesignation.TabIndex = 9
        '
        'lblCodeAgenceVoyager
        '
        Me.lblCodeAgenceVoyager.Location = New System.Drawing.Point(15, 99)
        Me.lblCodeAgenceVoyager.Name = "lblCodeAgenceVoyager"
        Me.lblCodeAgenceVoyager.Size = New System.Drawing.Size(165, 20)
        Me.lblCodeAgenceVoyager.TabIndex = 10
        Me.lblCodeAgenceVoyager.Text = "Code agence Voyager :"
        '
        'txtCodeAgenceVoyager
        '
        Me.txtCodeAgenceVoyager.Location = New System.Drawing.Point(185, 96)
        Me.txtCodeAgenceVoyager.MaxLength = 255
        Me.txtCodeAgenceVoyager.Name = "txtCodeAgenceVoyager"
        Me.txtCodeAgenceVoyager.Size = New System.Drawing.Size(280, 22)
        Me.txtCodeAgenceVoyager.TabIndex = 11
        '
        'btnNouveau
        '
        Me.btnNouveau.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnNouveau.Location = New System.Drawing.Point(12, 449)
        Me.btnNouveau.Name = "btnNouveau"
        Me.btnNouveau.Size = New System.Drawing.Size(120, 32)
        Me.btnNouveau.TabIndex = 12
        Me.btnNouveau.Text = "Nouveau"
        Me.btnNouveau.UseVisualStyleBackColor = True
        '
        'btnEnregistrer
        '
        Me.btnEnregistrer.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnEnregistrer.Location = New System.Drawing.Point(140, 449)
        Me.btnEnregistrer.Name = "btnEnregistrer"
        Me.btnEnregistrer.Size = New System.Drawing.Size(150, 32)
        Me.btnEnregistrer.TabIndex = 13
        Me.btnEnregistrer.Text = "Enregistrer"
        Me.btnEnregistrer.UseVisualStyleBackColor = True
        '
        'btnSupprimer
        '
        Me.btnSupprimer.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnSupprimer.Location = New System.Drawing.Point(298, 449)
        Me.btnSupprimer.Name = "btnSupprimer"
        Me.btnSupprimer.Size = New System.Drawing.Size(120, 32)
        Me.btnSupprimer.TabIndex = 14
        Me.btnSupprimer.Text = "Supprimer"
        Me.btnSupprimer.UseVisualStyleBackColor = True
        '
        'btnFermer
        '
        Me.btnFermer.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFermer.Location = New System.Drawing.Point(628, 449)
        Me.btnFermer.Name = "btnFermer"
        Me.btnFermer.Size = New System.Drawing.Size(120, 32)
        Me.btnFermer.TabIndex = 15
        Me.btnFermer.Text = "Fermer"
        Me.btnFermer.UseVisualStyleBackColor = True
        '
        'lblStatut
        '
        Me.lblStatut.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblStatut.ForeColor = System.Drawing.SystemColors.GrayText
        Me.lblStatut.Location = New System.Drawing.Point(12, 489)
        Me.lblStatut.Name = "lblStatut"
        Me.lblStatut.Size = New System.Drawing.Size(736, 20)
        Me.lblStatut.TabIndex = 16
        Me.lblStatut.Text = ""
        '
        'FrmAgences
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnFermer
        Me.ClientSize = New System.Drawing.Size(760, 519)
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
        Me.MinimumSize = New System.Drawing.Size(776, 559)
        Me.Name = "FrmAgences"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Agences propres Ecobank (T_Pdv_EC)"
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
    Friend WithEvents lblCodeSite As System.Windows.Forms.Label
    Friend WithEvents txtCodeSite As System.Windows.Forms.TextBox
    Friend WithEvents lblDesignation As System.Windows.Forms.Label
    Friend WithEvents txtDesignation As System.Windows.Forms.TextBox
    Friend WithEvents lblCodeAgenceVoyager As System.Windows.Forms.Label
    Friend WithEvents txtCodeAgenceVoyager As System.Windows.Forms.TextBox

End Class
