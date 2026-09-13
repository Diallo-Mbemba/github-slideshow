Option Strict On
Option Explicit On

Partial Class FrmChangerMotDePasse
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
        Me.panelTitre = New System.Windows.Forms.Panel()
        Me.lblTitre = New System.Windows.Forms.Label()
        Me.lblUtilisateur = New System.Windows.Forms.Label()
        Me.lblAncien = New System.Windows.Forms.Label()
        Me.txtAncien = New System.Windows.Forms.TextBox()
        Me.lblNouveau = New System.Windows.Forms.Label()
        Me.txtNouveau = New System.Windows.Forms.TextBox()
        Me.lblConfirmation = New System.Windows.Forms.Label()
        Me.txtConfirmation = New System.Windows.Forms.TextBox()
        Me.lblRegles = New System.Windows.Forms.Label()
        Me.lblMessage = New System.Windows.Forms.Label()
        Me.btnValider = New System.Windows.Forms.Button()
        Me.btnAnnuler = New System.Windows.Forms.Button()
        Me.panelTitre.SuspendLayout()
        Me.SuspendLayout()
        '
        'panelTitre
        '
        Me.panelTitre.BackColor = System.Drawing.Color.Black
        Me.panelTitre.Controls.Add(Me.lblTitre)
        Me.panelTitre.Location = New System.Drawing.Point(12, 12)
        Me.panelTitre.Name = "panelTitre"
        Me.panelTitre.Size = New System.Drawing.Size(446, 56)
        Me.panelTitre.TabIndex = 0
        '
        'lblTitre
        '
        Me.lblTitre.BackColor = System.Drawing.Color.Black
        Me.lblTitre.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblTitre.Font = New System.Drawing.Font("Segoe UI", 16.0!, System.Drawing.FontStyle.Bold)
        Me.lblTitre.ForeColor = System.Drawing.Color.Yellow
        Me.lblTitre.Location = New System.Drawing.Point(0, 0)
        Me.lblTitre.Name = "lblTitre"
        Me.lblTitre.Size = New System.Drawing.Size(446, 56)
        Me.lblTitre.TabIndex = 0
        Me.lblTitre.Text = "Changement de mot de passe"
        Me.lblTitre.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblUtilisateur
        '
        Me.lblUtilisateur.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold)
        Me.lblUtilisateur.Location = New System.Drawing.Point(12, 78)
        Me.lblUtilisateur.Name = "lblUtilisateur"
        Me.lblUtilisateur.Size = New System.Drawing.Size(446, 22)
        Me.lblUtilisateur.TabIndex = 1
        Me.lblUtilisateur.Text = ""
        Me.lblUtilisateur.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblAncien
        '
        Me.lblAncien.Location = New System.Drawing.Point(18, 114)
        Me.lblAncien.Name = "lblAncien"
        Me.lblAncien.Size = New System.Drawing.Size(152, 24)
        Me.lblAncien.TabIndex = 2
        Me.lblAncien.Text = "Mot de passe actuel"
        Me.lblAncien.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtAncien
        '
        Me.txtAncien.Location = New System.Drawing.Point(176, 112)
        Me.txtAncien.MaxLength = 128
        Me.txtAncien.Name = "txtAncien"
        Me.txtAncien.Size = New System.Drawing.Size(282, 22)
        Me.txtAncien.TabIndex = 3
        Me.txtAncien.UseSystemPasswordChar = True
        '
        'lblNouveau
        '
        Me.lblNouveau.Location = New System.Drawing.Point(18, 150)
        Me.lblNouveau.Name = "lblNouveau"
        Me.lblNouveau.Size = New System.Drawing.Size(152, 24)
        Me.lblNouveau.TabIndex = 4
        Me.lblNouveau.Text = "Nouveau mot de passe"
        Me.lblNouveau.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtNouveau
        '
        Me.txtNouveau.Location = New System.Drawing.Point(176, 148)
        Me.txtNouveau.MaxLength = 128
        Me.txtNouveau.Name = "txtNouveau"
        Me.txtNouveau.Size = New System.Drawing.Size(282, 22)
        Me.txtNouveau.TabIndex = 5
        Me.txtNouveau.UseSystemPasswordChar = True
        '
        'lblConfirmation
        '
        Me.lblConfirmation.Location = New System.Drawing.Point(18, 186)
        Me.lblConfirmation.Name = "lblConfirmation"
        Me.lblConfirmation.Size = New System.Drawing.Size(152, 24)
        Me.lblConfirmation.TabIndex = 6
        Me.lblConfirmation.Text = "Confirmation"
        Me.lblConfirmation.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtConfirmation
        '
        Me.txtConfirmation.Location = New System.Drawing.Point(176, 184)
        Me.txtConfirmation.MaxLength = 128
        Me.txtConfirmation.Name = "txtConfirmation"
        Me.txtConfirmation.Size = New System.Drawing.Size(282, 22)
        Me.txtConfirmation.TabIndex = 7
        Me.txtConfirmation.UseSystemPasswordChar = True
        '
        'lblRegles
        '
        Me.lblRegles.ForeColor = System.Drawing.Color.DimGray
        Me.lblRegles.Location = New System.Drawing.Point(18, 216)
        Me.lblRegles.Name = "lblRegles"
        Me.lblRegles.Size = New System.Drawing.Size(440, 46)
        Me.lblRegles.TabIndex = 8
        Me.lblRegles.Text = ""
        '
        'lblMessage
        '
        Me.lblMessage.ForeColor = System.Drawing.Color.Firebrick
        Me.lblMessage.Location = New System.Drawing.Point(18, 266)
        Me.lblMessage.Name = "lblMessage"
        Me.lblMessage.Size = New System.Drawing.Size(440, 44)
        Me.lblMessage.TabIndex = 9
        Me.lblMessage.Text = ""
        '
        'btnValider
        '
        Me.btnValider.Location = New System.Drawing.Point(176, 316)
        Me.btnValider.Name = "btnValider"
        Me.btnValider.Size = New System.Drawing.Size(136, 34)
        Me.btnValider.TabIndex = 10
        Me.btnValider.Text = "&Enregistrer"
        Me.btnValider.UseVisualStyleBackColor = True
        '
        'btnAnnuler
        '
        Me.btnAnnuler.Location = New System.Drawing.Point(322, 316)
        Me.btnAnnuler.Name = "btnAnnuler"
        Me.btnAnnuler.Size = New System.Drawing.Size(136, 34)
        Me.btnAnnuler.TabIndex = 11
        Me.btnAnnuler.Text = "&Annuler"
        Me.btnAnnuler.UseVisualStyleBackColor = True
        '
        'FrmChangerMotDePasse
        '
        Me.AcceptButton = Me.btnValider
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnAnnuler
        Me.ClientSize = New System.Drawing.Size(470, 364)
        Me.Controls.Add(Me.btnAnnuler)
        Me.Controls.Add(Me.btnValider)
        Me.Controls.Add(Me.lblMessage)
        Me.Controls.Add(Me.lblRegles)
        Me.Controls.Add(Me.txtConfirmation)
        Me.Controls.Add(Me.lblConfirmation)
        Me.Controls.Add(Me.txtNouveau)
        Me.Controls.Add(Me.lblNouveau)
        Me.Controls.Add(Me.txtAncien)
        Me.Controls.Add(Me.lblAncien)
        Me.Controls.Add(Me.lblUtilisateur)
        Me.Controls.Add(Me.panelTitre)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FrmChangerMotDePasse"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Mot de passe"
        Me.panelTitre.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()
    End Sub

    Friend WithEvents panelTitre As System.Windows.Forms.Panel
    Friend WithEvents lblTitre As System.Windows.Forms.Label
    Friend WithEvents lblUtilisateur As System.Windows.Forms.Label
    Friend WithEvents lblAncien As System.Windows.Forms.Label
    Friend WithEvents txtAncien As System.Windows.Forms.TextBox
    Friend WithEvents lblNouveau As System.Windows.Forms.Label
    Friend WithEvents txtNouveau As System.Windows.Forms.TextBox
    Friend WithEvents lblConfirmation As System.Windows.Forms.Label
    Friend WithEvents txtConfirmation As System.Windows.Forms.TextBox
    Friend WithEvents lblRegles As System.Windows.Forms.Label
    Friend WithEvents lblMessage As System.Windows.Forms.Label
    Friend WithEvents btnValider As System.Windows.Forms.Button
    Friend WithEvents btnAnnuler As System.Windows.Forms.Button

End Class
