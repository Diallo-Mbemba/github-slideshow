Option Strict On
Option Explicit On

Partial Class FrmConnexion
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
        Me.lblSousTitre = New System.Windows.Forms.Label()
        Me.lblIdentifiant = New System.Windows.Forms.Label()
        Me.txtIdentifiant = New System.Windows.Forms.TextBox()
        Me.lblMotDePasse = New System.Windows.Forms.Label()
        Me.txtMotDePasse = New System.Windows.Forms.TextBox()
        Me.lblMessage = New System.Windows.Forms.Label()
        Me.btnConnexion = New System.Windows.Forms.Button()
        Me.btnAnnuler = New System.Windows.Forms.Button()
        Me.lblBase = New System.Windows.Forms.Label()
        Me.panelTitre.SuspendLayout()
        Me.SuspendLayout()
        '
        'panelTitre
        '
        Me.panelTitre.BackColor = System.Drawing.Color.Black
        Me.panelTitre.Controls.Add(Me.lblTitre)
        Me.panelTitre.Location = New System.Drawing.Point(12, 12)
        Me.panelTitre.Name = "panelTitre"
        Me.panelTitre.Size = New System.Drawing.Size(420, 62)
        Me.panelTitre.TabIndex = 0
        '
        'lblTitre
        '
        Me.lblTitre.BackColor = System.Drawing.Color.Black
        Me.lblTitre.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblTitre.Font = New System.Drawing.Font("Segoe UI", 22.0!, System.Drawing.FontStyle.Bold)
        Me.lblTitre.ForeColor = System.Drawing.Color.Yellow
        Me.lblTitre.Location = New System.Drawing.Point(0, 0)
        Me.lblTitre.Name = "lblTitre"
        Me.lblTitre.Size = New System.Drawing.Size(420, 62)
        Me.lblTitre.TabIndex = 0
        Me.lblTitre.Text = "Wincompense TCHAD"
        Me.lblTitre.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblSousTitre
        '
        Me.lblSousTitre.Font = New System.Drawing.Font("Segoe UI", 10.0!, System.Drawing.FontStyle.Bold)
        Me.lblSousTitre.Location = New System.Drawing.Point(12, 84)
        Me.lblSousTitre.Name = "lblSousTitre"
        Me.lblSousTitre.Size = New System.Drawing.Size(420, 24)
        Me.lblSousTitre.TabIndex = 1
        Me.lblSousTitre.Text = "Identification de l'utilisateur"
        Me.lblSousTitre.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblIdentifiant
        '
        Me.lblIdentifiant.Location = New System.Drawing.Point(18, 120)
        Me.lblIdentifiant.Name = "lblIdentifiant"
        Me.lblIdentifiant.Size = New System.Drawing.Size(120, 24)
        Me.lblIdentifiant.TabIndex = 2
        Me.lblIdentifiant.Text = "Identifiant"
        Me.lblIdentifiant.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtIdentifiant
        '
        Me.txtIdentifiant.Location = New System.Drawing.Point(150, 118)
        Me.txtIdentifiant.MaxLength = 50
        Me.txtIdentifiant.Name = "txtIdentifiant"
        Me.txtIdentifiant.Size = New System.Drawing.Size(282, 22)
        Me.txtIdentifiant.TabIndex = 3
        '
        'lblMotDePasse
        '
        Me.lblMotDePasse.Location = New System.Drawing.Point(18, 156)
        Me.lblMotDePasse.Name = "lblMotDePasse"
        Me.lblMotDePasse.Size = New System.Drawing.Size(120, 24)
        Me.lblMotDePasse.TabIndex = 4
        Me.lblMotDePasse.Text = "Mot de passe"
        Me.lblMotDePasse.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtMotDePasse
        '
        Me.txtMotDePasse.Location = New System.Drawing.Point(150, 154)
        Me.txtMotDePasse.MaxLength = 128
        Me.txtMotDePasse.Name = "txtMotDePasse"
        Me.txtMotDePasse.Size = New System.Drawing.Size(282, 22)
        Me.txtMotDePasse.TabIndex = 5
        Me.txtMotDePasse.UseSystemPasswordChar = True
        '
        'lblMessage
        '
        Me.lblMessage.ForeColor = System.Drawing.Color.Firebrick
        Me.lblMessage.Location = New System.Drawing.Point(18, 188)
        Me.lblMessage.Name = "lblMessage"
        Me.lblMessage.Size = New System.Drawing.Size(414, 62)
        Me.lblMessage.TabIndex = 6
        Me.lblMessage.Text = ""
        '
        'btnConnexion
        '
        Me.btnConnexion.Location = New System.Drawing.Point(150, 258)
        Me.btnConnexion.Name = "btnConnexion"
        Me.btnConnexion.Size = New System.Drawing.Size(136, 34)
        Me.btnConnexion.TabIndex = 7
        Me.btnConnexion.Text = "&Se connecter"
        Me.btnConnexion.UseVisualStyleBackColor = True
        '
        'btnAnnuler
        '
        Me.btnAnnuler.Location = New System.Drawing.Point(296, 258)
        Me.btnAnnuler.Name = "btnAnnuler"
        Me.btnAnnuler.Size = New System.Drawing.Size(136, 34)
        Me.btnAnnuler.TabIndex = 8
        Me.btnAnnuler.Text = "&Quitter"
        Me.btnAnnuler.UseVisualStyleBackColor = True
        '
        'lblBase
        '
        Me.lblBase.ForeColor = System.Drawing.Color.DimGray
        Me.lblBase.Location = New System.Drawing.Point(12, 300)
        Me.lblBase.Name = "lblBase"
        Me.lblBase.Size = New System.Drawing.Size(420, 20)
        Me.lblBase.TabIndex = 9
        Me.lblBase.Text = ""
        Me.lblBase.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'FrmConnexion
        '
        Me.AcceptButton = Me.btnConnexion
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnAnnuler
        Me.ClientSize = New System.Drawing.Size(444, 330)
        Me.Controls.Add(Me.lblBase)
        Me.Controls.Add(Me.btnAnnuler)
        Me.Controls.Add(Me.btnConnexion)
        Me.Controls.Add(Me.lblMessage)
        Me.Controls.Add(Me.txtMotDePasse)
        Me.Controls.Add(Me.lblMotDePasse)
        Me.Controls.Add(Me.txtIdentifiant)
        Me.Controls.Add(Me.lblIdentifiant)
        Me.Controls.Add(Me.lblSousTitre)
        Me.Controls.Add(Me.panelTitre)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FrmConnexion"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Connexion — Wincompense TCHAD"
        Me.panelTitre.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()
    End Sub

    Friend WithEvents panelTitre As System.Windows.Forms.Panel
    Friend WithEvents lblTitre As System.Windows.Forms.Label
    Friend WithEvents lblSousTitre As System.Windows.Forms.Label
    Friend WithEvents lblIdentifiant As System.Windows.Forms.Label
    Friend WithEvents txtIdentifiant As System.Windows.Forms.TextBox
    Friend WithEvents lblMotDePasse As System.Windows.Forms.Label
    Friend WithEvents txtMotDePasse As System.Windows.Forms.TextBox
    Friend WithEvents lblMessage As System.Windows.Forms.Label
    Friend WithEvents btnConnexion As System.Windows.Forms.Button
    Friend WithEvents btnAnnuler As System.Windows.Forms.Button
    Friend WithEvents lblBase As System.Windows.Forms.Label

End Class
