Option Strict On
Option Explicit On

Partial Class FrmUtilisateurEdition
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
        Me.lblIdentifiant = New System.Windows.Forms.Label()
        Me.txtIdentifiant = New System.Windows.Forms.TextBox()
        Me.lblNomComplet = New System.Windows.Forms.Label()
        Me.txtNomComplet = New System.Windows.Forms.TextBox()
        Me.lblRole = New System.Windows.Forms.Label()
        Me.cboRole = New System.Windows.Forms.ComboBox()
        Me.lblDroits = New System.Windows.Forms.Label()
        Me.lblFonction = New System.Windows.Forms.Label()
        Me.cboFonction = New System.Windows.Forms.ComboBox()
        Me.lblAideFonction = New System.Windows.Forms.Label()
        Me.chkActif = New System.Windows.Forms.CheckBox()
        Me.lblMotDePasse = New System.Windows.Forms.Label()
        Me.txtMotDePasse = New System.Windows.Forms.TextBox()
        Me.lblConfirmation = New System.Windows.Forms.Label()
        Me.txtConfirmation = New System.Windows.Forms.TextBox()
        Me.lblRegles = New System.Windows.Forms.Label()
        Me.lblMessage = New System.Windows.Forms.Label()
        Me.btnEnregistrer = New System.Windows.Forms.Button()
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
        Me.panelTitre.Size = New System.Drawing.Size(496, 56)
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
        Me.lblTitre.Size = New System.Drawing.Size(496, 56)
        Me.lblTitre.TabIndex = 0
        Me.lblTitre.Text = "Compte utilisateur"
        Me.lblTitre.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblIdentifiant
        '
        Me.lblIdentifiant.Location = New System.Drawing.Point(18, 80)
        Me.lblIdentifiant.Name = "lblIdentifiant"
        Me.lblIdentifiant.Size = New System.Drawing.Size(152, 24)
        Me.lblIdentifiant.TabIndex = 1
        Me.lblIdentifiant.Text = "Identifiant"
        Me.lblIdentifiant.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtIdentifiant
        '
        Me.txtIdentifiant.Location = New System.Drawing.Point(176, 78)
        Me.txtIdentifiant.MaxLength = 50
        Me.txtIdentifiant.Name = "txtIdentifiant"
        Me.txtIdentifiant.Size = New System.Drawing.Size(332, 22)
        Me.txtIdentifiant.TabIndex = 2
        '
        'lblNomComplet
        '
        Me.lblNomComplet.Location = New System.Drawing.Point(18, 116)
        Me.lblNomComplet.Name = "lblNomComplet"
        Me.lblNomComplet.Size = New System.Drawing.Size(152, 24)
        Me.lblNomComplet.TabIndex = 3
        Me.lblNomComplet.Text = "Nom et prénom"
        Me.lblNomComplet.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtNomComplet
        '
        Me.txtNomComplet.Location = New System.Drawing.Point(176, 114)
        Me.txtNomComplet.MaxLength = 150
        Me.txtNomComplet.Name = "txtNomComplet"
        Me.txtNomComplet.Size = New System.Drawing.Size(332, 22)
        Me.txtNomComplet.TabIndex = 4
        '
        'lblRole
        '
        Me.lblRole.Location = New System.Drawing.Point(18, 152)
        Me.lblRole.Name = "lblRole"
        Me.lblRole.Size = New System.Drawing.Size(152, 24)
        Me.lblRole.TabIndex = 5
        Me.lblRole.Text = "Rôle"
        Me.lblRole.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'cboRole
        '
        Me.cboRole.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboRole.FormattingEnabled = True
        Me.cboRole.Location = New System.Drawing.Point(176, 150)
        Me.cboRole.Name = "cboRole"
        Me.cboRole.Size = New System.Drawing.Size(332, 24)
        Me.cboRole.TabIndex = 6
        '
        'lblDroits
        '
        Me.lblDroits.ForeColor = System.Drawing.Color.DimGray
        Me.lblDroits.Location = New System.Drawing.Point(176, 180)
        Me.lblDroits.Name = "lblDroits"
        Me.lblDroits.Size = New System.Drawing.Size(332, 58)
        Me.lblDroits.TabIndex = 7
        Me.lblDroits.Text = ""
        '
        'lblFonction
        '
        Me.lblFonction.Location = New System.Drawing.Point(18, 244)
        Me.lblFonction.Name = "lblFonction"
        Me.lblFonction.Size = New System.Drawing.Size(152, 24)
        Me.lblFonction.TabIndex = 20
        Me.lblFonction.Text = "Fonction (référentiel)"
        Me.lblFonction.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'cboFonction
        '
        Me.cboFonction.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboFonction.FormattingEnabled = True
        Me.cboFonction.Location = New System.Drawing.Point(176, 242)
        Me.cboFonction.Name = "cboFonction"
        Me.cboFonction.Size = New System.Drawing.Size(332, 24)
        Me.cboFonction.TabIndex = 21
        '
        'lblAideFonction
        '
        Me.lblAideFonction.ForeColor = System.Drawing.SystemColors.GrayText
        Me.lblAideFonction.Location = New System.Drawing.Point(176, 268)
        Me.lblAideFonction.Name = "lblAideFonction"
        Me.lblAideFonction.Size = New System.Drawing.Size(332, 20)
        Me.lblAideFonction.TabIndex = 22
        Me.lblAideFonction.Text = "Facultative : elle s'attribue aussi plus tard."
        '
        'chkActif
        '
        Me.chkActif.Checked = True
        Me.chkActif.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkActif.Location = New System.Drawing.Point(176, 302)
        Me.chkActif.Name = "chkActif"
        Me.chkActif.Size = New System.Drawing.Size(332, 24)
        Me.chkActif.TabIndex = 8
        Me.chkActif.Text = "Compte actif (peut se connecter)"
        Me.chkActif.UseVisualStyleBackColor = True
        '
        'lblMotDePasse
        '
        Me.lblMotDePasse.Location = New System.Drawing.Point(18, 336)
        Me.lblMotDePasse.Name = "lblMotDePasse"
        Me.lblMotDePasse.Size = New System.Drawing.Size(152, 24)
        Me.lblMotDePasse.TabIndex = 9
        Me.lblMotDePasse.Text = "Mot de passe initial"
        Me.lblMotDePasse.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtMotDePasse
        '
        Me.txtMotDePasse.Location = New System.Drawing.Point(176, 334)
        Me.txtMotDePasse.MaxLength = 128
        Me.txtMotDePasse.Name = "txtMotDePasse"
        Me.txtMotDePasse.Size = New System.Drawing.Size(332, 22)
        Me.txtMotDePasse.TabIndex = 10
        Me.txtMotDePasse.UseSystemPasswordChar = True
        '
        'lblConfirmation
        '
        Me.lblConfirmation.Location = New System.Drawing.Point(18, 372)
        Me.lblConfirmation.Name = "lblConfirmation"
        Me.lblConfirmation.Size = New System.Drawing.Size(152, 24)
        Me.lblConfirmation.TabIndex = 11
        Me.lblConfirmation.Text = "Confirmation"
        Me.lblConfirmation.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtConfirmation
        '
        Me.txtConfirmation.Location = New System.Drawing.Point(176, 370)
        Me.txtConfirmation.MaxLength = 128
        Me.txtConfirmation.Name = "txtConfirmation"
        Me.txtConfirmation.Size = New System.Drawing.Size(332, 22)
        Me.txtConfirmation.TabIndex = 12
        Me.txtConfirmation.UseSystemPasswordChar = True
        '
        'lblRegles
        '
        Me.lblRegles.ForeColor = System.Drawing.Color.DimGray
        Me.lblRegles.Location = New System.Drawing.Point(18, 400)
        Me.lblRegles.Name = "lblRegles"
        Me.lblRegles.Size = New System.Drawing.Size(490, 44)
        Me.lblRegles.TabIndex = 13
        Me.lblRegles.Text = ""
        '
        'lblMessage
        '
        Me.lblMessage.ForeColor = System.Drawing.Color.Firebrick
        Me.lblMessage.Location = New System.Drawing.Point(18, 446)
        Me.lblMessage.Name = "lblMessage"
        Me.lblMessage.Size = New System.Drawing.Size(490, 32)
        Me.lblMessage.TabIndex = 14
        Me.lblMessage.Text = ""
        '
        'btnEnregistrer
        '
        Me.btnEnregistrer.Location = New System.Drawing.Point(232, 482)
        Me.btnEnregistrer.Name = "btnEnregistrer"
        Me.btnEnregistrer.Size = New System.Drawing.Size(134, 34)
        Me.btnEnregistrer.TabIndex = 15
        Me.btnEnregistrer.Text = "&Enregistrer"
        Me.btnEnregistrer.UseVisualStyleBackColor = True
        '
        'btnAnnuler
        '
        Me.btnAnnuler.Location = New System.Drawing.Point(374, 482)
        Me.btnAnnuler.Name = "btnAnnuler"
        Me.btnAnnuler.Size = New System.Drawing.Size(134, 34)
        Me.btnAnnuler.TabIndex = 16
        Me.btnAnnuler.Text = "&Annuler"
        Me.btnAnnuler.UseVisualStyleBackColor = True
        '
        'FrmUtilisateurEdition
        '
        Me.AcceptButton = Me.btnEnregistrer
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnAnnuler
        Me.ClientSize = New System.Drawing.Size(520, 528)
        Me.Controls.Add(Me.btnAnnuler)
        Me.Controls.Add(Me.btnEnregistrer)
        Me.Controls.Add(Me.lblMessage)
        Me.Controls.Add(Me.lblRegles)
        Me.Controls.Add(Me.txtConfirmation)
        Me.Controls.Add(Me.lblConfirmation)
        Me.Controls.Add(Me.txtMotDePasse)
        Me.Controls.Add(Me.lblMotDePasse)
        Me.Controls.Add(Me.chkActif)
        Me.Controls.Add(Me.lblAideFonction)
        Me.Controls.Add(Me.cboFonction)
        Me.Controls.Add(Me.lblFonction)
        Me.Controls.Add(Me.lblDroits)
        Me.Controls.Add(Me.cboRole)
        Me.Controls.Add(Me.lblRole)
        Me.Controls.Add(Me.txtNomComplet)
        Me.Controls.Add(Me.lblNomComplet)
        Me.Controls.Add(Me.txtIdentifiant)
        Me.Controls.Add(Me.lblIdentifiant)
        Me.Controls.Add(Me.panelTitre)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FrmUtilisateurEdition"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Compte utilisateur"
        Me.panelTitre.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()
    End Sub

    Friend WithEvents panelTitre As System.Windows.Forms.Panel
    Friend WithEvents lblTitre As System.Windows.Forms.Label
    Friend WithEvents lblIdentifiant As System.Windows.Forms.Label
    Friend WithEvents txtIdentifiant As System.Windows.Forms.TextBox
    Friend WithEvents lblNomComplet As System.Windows.Forms.Label
    Friend WithEvents txtNomComplet As System.Windows.Forms.TextBox
    Friend WithEvents lblRole As System.Windows.Forms.Label
    Friend WithEvents cboRole As System.Windows.Forms.ComboBox
    Friend WithEvents lblDroits As System.Windows.Forms.Label
    Friend WithEvents lblFonction As System.Windows.Forms.Label
    Friend WithEvents cboFonction As System.Windows.Forms.ComboBox
    Friend WithEvents lblAideFonction As System.Windows.Forms.Label
    Friend WithEvents chkActif As System.Windows.Forms.CheckBox
    Friend WithEvents lblMotDePasse As System.Windows.Forms.Label
    Friend WithEvents txtMotDePasse As System.Windows.Forms.TextBox
    Friend WithEvents lblConfirmation As System.Windows.Forms.Label
    Friend WithEvents txtConfirmation As System.Windows.Forms.TextBox
    Friend WithEvents lblRegles As System.Windows.Forms.Label
    Friend WithEvents lblMessage As System.Windows.Forms.Label
    Friend WithEvents btnEnregistrer As System.Windows.Forms.Button
    Friend WithEvents btnAnnuler As System.Windows.Forms.Button

End Class
