Option Strict On
Option Explicit On

Partial Class FrmParametresConnexion
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
        Me.lblOrigine = New System.Windows.Forms.Label()
        Me.grpServeur = New System.Windows.Forms.GroupBox()
        Me.lblServeur = New System.Windows.Forms.Label()
        Me.txtServeur = New System.Windows.Forms.TextBox()
        Me.lblBase = New System.Windows.Forms.Label()
        Me.txtBase = New System.Windows.Forms.TextBox()
        Me.lblDelai = New System.Windows.Forms.Label()
        Me.nudDelai = New System.Windows.Forms.NumericUpDown()
        Me.lblSecondes = New System.Windows.Forms.Label()
        Me.grpChaine = New System.Windows.Forms.GroupBox()
        Me.chkChaineComplete = New System.Windows.Forms.CheckBox()
        Me.lblAideChaine = New System.Windows.Forms.Label()
        Me.txtChaine = New System.Windows.Forms.TextBox()
        Me.grpPartage = New System.Windows.Forms.GroupBox()
        Me.chkConserver = New System.Windows.Forms.CheckBox()
        Me.lblAideConserver = New System.Windows.Forms.Label()
        Me.lblPartage = New System.Windows.Forms.Label()
        Me.txtPartage = New System.Windows.Forms.TextBox()
        Me.chkPropager = New System.Windows.Forms.CheckBox()
        Me.lblAidePartage = New System.Windows.Forms.Label()
        Me.btnTester = New System.Windows.Forms.Button()
        Me.btnPreparer = New System.Windows.Forms.Button()
        Me.lblResultat = New System.Windows.Forms.Label()
        Me.btnEnregistrer = New System.Windows.Forms.Button()
        Me.btnFermer = New System.Windows.Forms.Button()
        CType(Me.nudDelai, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpServeur.SuspendLayout()
        Me.grpChaine.SuspendLayout()
        Me.grpPartage.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblTitre
        '
        Me.lblTitre.AutoEllipsis = True
        Me.lblTitre.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.0!, System.Drawing.FontStyle.Bold)
        Me.lblTitre.Location = New System.Drawing.Point(12, 12)
        Me.lblTitre.Name = "lblTitre"
        Me.lblTitre.Size = New System.Drawing.Size(560, 22)
        Me.lblTitre.TabIndex = 0
        Me.lblTitre.Text = "Connexion à la base de données"
        '
        'lblOrigine
        '
        Me.lblOrigine.AutoEllipsis = True
        Me.lblOrigine.ForeColor = System.Drawing.SystemColors.ControlDarkDark
        Me.lblOrigine.Location = New System.Drawing.Point(12, 36)
        Me.lblOrigine.Name = "lblOrigine"
        Me.lblOrigine.Size = New System.Drawing.Size(560, 32)
        Me.lblOrigine.TabIndex = 1
        '
        'grpServeur
        '
        Me.grpServeur.Controls.Add(Me.lblServeur)
        Me.grpServeur.Controls.Add(Me.txtServeur)
        Me.grpServeur.Controls.Add(Me.lblBase)
        Me.grpServeur.Controls.Add(Me.txtBase)
        Me.grpServeur.Controls.Add(Me.lblDelai)
        Me.grpServeur.Controls.Add(Me.nudDelai)
        Me.grpServeur.Controls.Add(Me.lblSecondes)
        Me.grpServeur.Location = New System.Drawing.Point(12, 74)
        Me.grpServeur.Name = "grpServeur"
        Me.grpServeur.Size = New System.Drawing.Size(560, 124)
        Me.grpServeur.TabIndex = 2
        Me.grpServeur.TabStop = False
        Me.grpServeur.Text = "Serveur SQL Server"
        '
        'lblServeur
        '
        Me.lblServeur.Location = New System.Drawing.Point(14, 28)
        Me.lblServeur.Name = "lblServeur"
        Me.lblServeur.Size = New System.Drawing.Size(120, 20)
        Me.lblServeur.TabIndex = 0
        Me.lblServeur.Text = "Serveur ou instance"
        '
        'txtServeur
        '
        Me.txtServeur.Location = New System.Drawing.Point(140, 25)
        Me.txtServeur.MaxLength = 200
        Me.txtServeur.Name = "txtServeur"
        Me.txtServeur.Size = New System.Drawing.Size(400, 20)
        Me.txtServeur.TabIndex = 1
        '
        'lblBase
        '
        Me.lblBase.Location = New System.Drawing.Point(14, 58)
        Me.lblBase.Name = "lblBase"
        Me.lblBase.Size = New System.Drawing.Size(120, 20)
        Me.lblBase.TabIndex = 2
        Me.lblBase.Text = "Base de données"
        '
        'txtBase
        '
        Me.txtBase.Location = New System.Drawing.Point(140, 55)
        Me.txtBase.MaxLength = 128
        Me.txtBase.Name = "txtBase"
        Me.txtBase.Size = New System.Drawing.Size(400, 20)
        Me.txtBase.TabIndex = 3
        '
        'lblDelai
        '
        Me.lblDelai.Location = New System.Drawing.Point(14, 90)
        Me.lblDelai.Name = "lblDelai"
        Me.lblDelai.Size = New System.Drawing.Size(120, 20)
        Me.lblDelai.TabIndex = 4
        Me.lblDelai.Text = "Délai de connexion"
        '
        'nudDelai
        '
        Me.nudDelai.Location = New System.Drawing.Point(140, 88)
        Me.nudDelai.Maximum = New Decimal(New Integer() {120, 0, 0, 0})
        Me.nudDelai.Minimum = New Decimal(New Integer() {3, 0, 0, 0})
        Me.nudDelai.Name = "nudDelai"
        Me.nudDelai.Size = New System.Drawing.Size(60, 20)
        Me.nudDelai.TabIndex = 5
        Me.nudDelai.Value = New Decimal(New Integer() {10, 0, 0, 0})
        '
        'lblSecondes
        '
        Me.lblSecondes.ForeColor = System.Drawing.SystemColors.ControlDarkDark
        Me.lblSecondes.Location = New System.Drawing.Point(206, 90)
        Me.lblSecondes.Name = "lblSecondes"
        Me.lblSecondes.Size = New System.Drawing.Size(334, 20)
        Me.lblSecondes.TabIndex = 6
        Me.lblSecondes.Text = "secondes — authentification Windows intégrée, aucun mot de passe."
        '
        'grpChaine
        '
        Me.grpChaine.Controls.Add(Me.chkChaineComplete)
        Me.grpChaine.Controls.Add(Me.lblAideChaine)
        Me.grpChaine.Controls.Add(Me.txtChaine)
        Me.grpChaine.Location = New System.Drawing.Point(12, 206)
        Me.grpChaine.Name = "grpChaine"
        Me.grpChaine.Size = New System.Drawing.Size(560, 104)
        Me.grpChaine.TabIndex = 3
        Me.grpChaine.TabStop = False
        Me.grpChaine.Text = "Chaîne fournie par la banque"
        '
        'chkChaineComplete
        '
        Me.chkChaineComplete.Location = New System.Drawing.Point(14, 24)
        Me.chkChaineComplete.Name = "chkChaineComplete"
        Me.chkChaineComplete.Size = New System.Drawing.Size(530, 20)
        Me.chkChaineComplete.TabIndex = 0
        Me.chkChaineComplete.Text = "Employer une chaîne de connexion complète, telle qu'elle a été fournie"
        Me.chkChaineComplete.UseVisualStyleBackColor = True
        '
        'lblAideChaine
        '
        Me.lblAideChaine.ForeColor = System.Drawing.SystemColors.ControlDarkDark
        Me.lblAideChaine.Location = New System.Drawing.Point(32, 46)
        Me.lblAideChaine.Name = "lblAideChaine"
        Me.lblAideChaine.Size = New System.Drawing.Size(512, 18)
        Me.lblAideChaine.TabIndex = 1
        Me.lblAideChaine.Text = "Elle remplace le serveur et la base ci-dessus. Aucun mot de passe ne doit y figure" &
            "r."
        '
        'txtChaine
        '
        Me.txtChaine.Enabled = False
        Me.txtChaine.Location = New System.Drawing.Point(14, 68)
        Me.txtChaine.MaxLength = 1000
        Me.txtChaine.Name = "txtChaine"
        Me.txtChaine.Size = New System.Drawing.Size(530, 20)
        Me.txtChaine.TabIndex = 2
        '
        'grpPartage
        '
        Me.grpPartage.Controls.Add(Me.chkConserver)
        Me.grpPartage.Controls.Add(Me.lblAideConserver)
        Me.grpPartage.Controls.Add(Me.lblPartage)
        Me.grpPartage.Controls.Add(Me.txtPartage)
        Me.grpPartage.Controls.Add(Me.chkPropager)
        Me.grpPartage.Controls.Add(Me.lblAidePartage)
        Me.grpPartage.Location = New System.Drawing.Point(12, 318)
        Me.grpPartage.Name = "grpPartage"
        Me.grpPartage.Size = New System.Drawing.Size(560, 166)
        Me.grpPartage.TabIndex = 4
        Me.grpPartage.TabStop = False
        Me.grpPartage.Text = "Portée du réglage"
        '
        'chkConserver
        '
        Me.chkConserver.Location = New System.Drawing.Point(14, 24)
        Me.chkConserver.Name = "chkConserver"
        Me.chkConserver.Size = New System.Drawing.Size(530, 20)
        Me.chkConserver.TabIndex = 0
        Me.chkConserver.Text = "Conserver ce réglage sur ce poste"
        Me.chkConserver.UseVisualStyleBackColor = True
        '
        'lblAideConserver
        '
        Me.lblAideConserver.ForeColor = System.Drawing.SystemColors.ControlDarkDark
        Me.lblAideConserver.Location = New System.Drawing.Point(32, 46)
        Me.lblAideConserver.Name = "lblAideConserver"
        Me.lblAideConserver.Size = New System.Drawing.Size(512, 32)
        Me.lblAideConserver.TabIndex = 1
        Me.lblAideConserver.Text = "Sans cette case, le réglage ne vaut que pour cette session : rien n'est écrit sur" &
            " le poste, qui continue de suivre la chaîne publiée avec l'application."
        '
        'lblPartage
        '
        Me.lblPartage.Location = New System.Drawing.Point(14, 86)
        Me.lblPartage.Name = "lblPartage"
        Me.lblPartage.Size = New System.Drawing.Size(120, 20)
        Me.lblPartage.TabIndex = 2
        Me.lblPartage.Text = "Chemin du fichier"
        '
        'txtPartage
        '
        Me.txtPartage.Location = New System.Drawing.Point(140, 83)
        Me.txtPartage.MaxLength = 400
        Me.txtPartage.Name = "txtPartage"
        Me.txtPartage.Size = New System.Drawing.Size(400, 20)
        Me.txtPartage.TabIndex = 3
        '
        'chkPropager
        '
        Me.chkPropager.Location = New System.Drawing.Point(140, 110)
        Me.chkPropager.Name = "chkPropager"
        Me.chkPropager.Size = New System.Drawing.Size(400, 20)
        Me.chkPropager.TabIndex = 4
        Me.chkPropager.Text = "Appliquer ce réglage à TOUS les postes de la banque"
        Me.chkPropager.UseVisualStyleBackColor = True
        '
        'lblAidePartage
        '
        Me.lblAidePartage.ForeColor = System.Drawing.SystemColors.ControlDarkDark
        Me.lblAidePartage.Location = New System.Drawing.Point(140, 132)
        Me.lblAidePartage.Name = "lblAidePartage"
        Me.lblAidePartage.Size = New System.Drawing.Size(400, 28)
        Me.lblAidePartage.TabIndex = 5
        Me.lblAidePartage.Text = "Sans cette case, le réglage ne vaut que pour ce poste. Les autres continueront de" &
            " lire l'ancien serveur."
        '
        'btnTester
        '
        Me.btnTester.Location = New System.Drawing.Point(12, 498)
        Me.btnTester.Name = "btnTester"
        Me.btnTester.Size = New System.Drawing.Size(150, 32)
        Me.btnTester.TabIndex = 5
        Me.btnTester.Text = "Tester la connexion"
        Me.btnTester.UseVisualStyleBackColor = True
        '
        'lblResultat
        '
        Me.lblResultat.Location = New System.Drawing.Point(174, 494)
        Me.lblResultat.Name = "lblResultat"
        Me.lblResultat.Size = New System.Drawing.Size(398, 40)
        Me.lblResultat.TabIndex = 6
        '
        'btnPreparer
        '
        Me.btnPreparer.Location = New System.Drawing.Point(12, 546)
        Me.btnPreparer.Name = "btnPreparer"
        Me.btnPreparer.Size = New System.Drawing.Size(190, 32)
        Me.btnPreparer.TabIndex = 7
        Me.btnPreparer.Text = "Préparer la base…"
        Me.btnPreparer.UseVisualStyleBackColor = True
        '
        'btnEnregistrer
        '
        Me.btnEnregistrer.Location = New System.Drawing.Point(352, 546)
        Me.btnEnregistrer.Name = "btnEnregistrer"
        Me.btnEnregistrer.Size = New System.Drawing.Size(120, 32)
        Me.btnEnregistrer.TabIndex = 8
        Me.btnEnregistrer.Text = "Enregistrer"
        Me.btnEnregistrer.UseVisualStyleBackColor = True
        '
        'btnFermer
        '
        Me.btnFermer.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnFermer.Location = New System.Drawing.Point(478, 546)
        Me.btnFermer.Name = "btnFermer"
        Me.btnFermer.Size = New System.Drawing.Size(94, 32)
        Me.btnFermer.TabIndex = 9
        Me.btnFermer.Text = "Fermer"
        Me.btnFermer.UseVisualStyleBackColor = True
        '
        'FrmParametresConnexion
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnFermer
        Me.ClientSize = New System.Drawing.Size(584, 590)
        Me.Controls.Add(Me.btnFermer)
        Me.Controls.Add(Me.btnEnregistrer)
        Me.Controls.Add(Me.lblResultat)
        Me.Controls.Add(Me.btnPreparer)
        Me.Controls.Add(Me.btnTester)
        Me.Controls.Add(Me.grpPartage)
        Me.Controls.Add(Me.grpChaine)
        Me.Controls.Add(Me.grpServeur)
        Me.Controls.Add(Me.lblOrigine)
        Me.Controls.Add(Me.lblTitre)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FrmParametresConnexion"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Connexion à la base de données"
        CType(Me.nudDelai, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpServeur.ResumeLayout(False)
        Me.grpServeur.PerformLayout()
        Me.grpChaine.ResumeLayout(False)
        Me.grpChaine.PerformLayout()
        Me.grpPartage.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents lblTitre As System.Windows.Forms.Label
    Friend WithEvents lblOrigine As System.Windows.Forms.Label
    Friend WithEvents grpServeur As System.Windows.Forms.GroupBox
    Friend WithEvents lblServeur As System.Windows.Forms.Label
    Friend WithEvents txtServeur As System.Windows.Forms.TextBox
    Friend WithEvents lblBase As System.Windows.Forms.Label
    Friend WithEvents txtBase As System.Windows.Forms.TextBox
    Friend WithEvents lblDelai As System.Windows.Forms.Label
    Friend WithEvents nudDelai As System.Windows.Forms.NumericUpDown
    Friend WithEvents lblSecondes As System.Windows.Forms.Label
    Friend WithEvents grpChaine As System.Windows.Forms.GroupBox
    Friend WithEvents chkChaineComplete As System.Windows.Forms.CheckBox
    Friend WithEvents lblAideChaine As System.Windows.Forms.Label
    Friend WithEvents txtChaine As System.Windows.Forms.TextBox
    Friend WithEvents grpPartage As System.Windows.Forms.GroupBox
    Friend WithEvents chkConserver As System.Windows.Forms.CheckBox
    Friend WithEvents lblAideConserver As System.Windows.Forms.Label
    Friend WithEvents lblPartage As System.Windows.Forms.Label
    Friend WithEvents txtPartage As System.Windows.Forms.TextBox
    Friend WithEvents chkPropager As System.Windows.Forms.CheckBox
    Friend WithEvents lblAidePartage As System.Windows.Forms.Label
    Friend WithEvents btnTester As System.Windows.Forms.Button
    Friend WithEvents btnPreparer As System.Windows.Forms.Button
    Friend WithEvents lblResultat As System.Windows.Forms.Label
    Friend WithEvents btnEnregistrer As System.Windows.Forms.Button
    Friend WithEvents btnFermer As System.Windows.Forms.Button

End Class
