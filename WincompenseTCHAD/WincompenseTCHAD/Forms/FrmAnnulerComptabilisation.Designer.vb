Option Strict On
Option Explicit On

Partial Class FrmAnnulerComptabilisation
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
        Me.lblJournee = New System.Windows.Forms.Label()
        Me.lblContenu = New System.Windows.Forms.Label()
        Me.grpCoreBanking = New System.Windows.Forms.GroupBox()
        Me.lblQuestion = New System.Windows.Forms.Label()
        Me.rdoNonInjecte = New System.Windows.Forms.RadioButton()
        Me.rdoInjecte = New System.Windows.Forms.RadioButton()
        Me.lblExtourne = New System.Windows.Forms.Label()
        Me.lblMotif = New System.Windows.Forms.Label()
        Me.cboMotif = New System.Windows.Forms.ComboBox()
        Me.lblCommentaire = New System.Windows.Forms.Label()
        Me.txtCommentaire = New System.Windows.Forms.TextBox()
        Me.lblAvertissement = New System.Windows.Forms.Label()
        Me.btnDeposer = New System.Windows.Forms.Button()
        Me.btnRenoncer = New System.Windows.Forms.Button()
        Me.grpCoreBanking.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblTitre
        '
        Me.lblTitre.BackColor = System.Drawing.Color.Black
        Me.lblTitre.Font = New System.Drawing.Font("Segoe UI", 14.0!, System.Drawing.FontStyle.Bold)
        Me.lblTitre.ForeColor = System.Drawing.Color.Yellow
        Me.lblTitre.Location = New System.Drawing.Point(12, 12)
        Me.lblTitre.Name = "lblTitre"
        Me.lblTitre.Size = New System.Drawing.Size(636, 34)
        Me.lblTitre.TabIndex = 0
        Me.lblTitre.Text = "   Demander l'annulation d'une comptabilisation"
        Me.lblTitre.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblJournee
        '
        Me.lblJournee.Font = New System.Drawing.Font("Segoe UI", 10.0!, System.Drawing.FontStyle.Bold)
        Me.lblJournee.Location = New System.Drawing.Point(12, 54)
        Me.lblJournee.Name = "lblJournee"
        Me.lblJournee.Size = New System.Drawing.Size(636, 24)
        Me.lblJournee.TabIndex = 1
        Me.lblJournee.Text = "Journée du"
        '
        'lblContenu
        '
        Me.lblContenu.ForeColor = System.Drawing.SystemColors.GrayText
        Me.lblContenu.Location = New System.Drawing.Point(12, 80)
        Me.lblContenu.Name = "lblContenu"
        Me.lblContenu.Size = New System.Drawing.Size(636, 40)
        Me.lblContenu.TabIndex = 2
        '
        'grpCoreBanking
        '
        Me.grpCoreBanking.Controls.Add(Me.lblQuestion)
        Me.grpCoreBanking.Controls.Add(Me.rdoNonInjecte)
        Me.grpCoreBanking.Controls.Add(Me.rdoInjecte)
        Me.grpCoreBanking.Location = New System.Drawing.Point(12, 124)
        Me.grpCoreBanking.Name = "grpCoreBanking"
        Me.grpCoreBanking.Size = New System.Drawing.Size(636, 120)
        Me.grpCoreBanking.TabIndex = 3
        Me.grpCoreBanking.TabStop = False
        Me.grpCoreBanking.Text = "Fichier destiné au core banking"
        '
        'lblQuestion
        '
        Me.lblQuestion.Location = New System.Drawing.Point(12, 22)
        Me.lblQuestion.Name = "lblQuestion"
        Me.lblQuestion.Size = New System.Drawing.Size(612, 36)
        Me.lblQuestion.TabIndex = 0
        Me.lblQuestion.Text = "Ce fichier a-t-il déjà été injecté dans le core banking ?"
        '
        'rdoNonInjecte
        '
        Me.rdoNonInjecte.Checked = True
        Me.rdoNonInjecte.Location = New System.Drawing.Point(24, 60)
        Me.rdoNonInjecte.Name = "rdoNonInjecte"
        Me.rdoNonInjecte.Size = New System.Drawing.Size(600, 24)
        Me.rdoNonInjecte.TabIndex = 1
        Me.rdoNonInjecte.TabStop = True
        Me.rdoNonInjecte.Text = "Non — aucune écriture n'est partie en comptabilité"
        Me.rdoNonInjecte.UseVisualStyleBackColor = True
        '
        'rdoInjecte
        '
        Me.rdoInjecte.Location = New System.Drawing.Point(24, 86)
        Me.rdoInjecte.Name = "rdoInjecte"
        Me.rdoInjecte.Size = New System.Drawing.Size(600, 24)
        Me.rdoInjecte.TabIndex = 2
        Me.rdoInjecte.Text = "Oui — les écritures sont dans les livres de la banque"
        Me.rdoInjecte.UseVisualStyleBackColor = True
        '
        'lblExtourne
        '
        Me.lblExtourne.ForeColor = System.Drawing.Color.Firebrick
        Me.lblExtourne.Location = New System.Drawing.Point(12, 250)
        Me.lblExtourne.Name = "lblExtourne"
        Me.lblExtourne.Size = New System.Drawing.Size(636, 44)
        Me.lblExtourne.TabIndex = 4
        '
        'lblMotif
        '
        Me.lblMotif.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold)
        Me.lblMotif.Location = New System.Drawing.Point(12, 300)
        Me.lblMotif.Name = "lblMotif"
        Me.lblMotif.Size = New System.Drawing.Size(300, 20)
        Me.lblMotif.TabIndex = 5
        Me.lblMotif.Text = "Motif de l'annulation (obligatoire)"
        '
        'cboMotif
        '
        Me.cboMotif.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboMotif.FormattingEnabled = True
        Me.cboMotif.Location = New System.Drawing.Point(12, 322)
        Me.cboMotif.Name = "cboMotif"
        Me.cboMotif.Size = New System.Drawing.Size(420, 21)
        Me.cboMotif.TabIndex = 6
        '
        'lblCommentaire
        '
        Me.lblCommentaire.Location = New System.Drawing.Point(12, 354)
        Me.lblCommentaire.Name = "lblCommentaire"
        Me.lblCommentaire.Size = New System.Drawing.Size(636, 20)
        Me.lblCommentaire.TabIndex = 7
        Me.lblCommentaire.Text = "Explication (conservée dans l'archive)"
        '
        'txtCommentaire
        '
        Me.txtCommentaire.Location = New System.Drawing.Point(12, 376)
        Me.txtCommentaire.MaxLength = 500
        Me.txtCommentaire.Multiline = True
        Me.txtCommentaire.Name = "txtCommentaire"
        Me.txtCommentaire.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtCommentaire.Size = New System.Drawing.Size(636, 100)
        Me.txtCommentaire.TabIndex = 8
        '
        'lblAvertissement
        '
        Me.lblAvertissement.ForeColor = System.Drawing.SystemColors.GrayText
        Me.lblAvertissement.Location = New System.Drawing.Point(12, 484)
        Me.lblAvertissement.Name = "lblAvertissement"
        Me.lblAvertissement.Size = New System.Drawing.Size(636, 46)
        Me.lblAvertissement.TabIndex = 9
        Me.lblAvertissement.Text = "Cette demande n'a aucun effet tant qu'un authorizer ne l'a pas autorisée : la jour" &
            "née reste dans les rapports jusque-là. Vous ne pourrez pas la décider vous-même."
        '
        'btnDeposer
        '
        Me.btnDeposer.Location = New System.Drawing.Point(368, 540)
        Me.btnDeposer.Name = "btnDeposer"
        Me.btnDeposer.Size = New System.Drawing.Size(170, 32)
        Me.btnDeposer.TabIndex = 10
        Me.btnDeposer.Text = "Déposer la demande"
        Me.btnDeposer.UseVisualStyleBackColor = True
        '
        'btnRenoncer
        '
        Me.btnRenoncer.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnRenoncer.Location = New System.Drawing.Point(548, 540)
        Me.btnRenoncer.Name = "btnRenoncer"
        Me.btnRenoncer.Size = New System.Drawing.Size(100, 32)
        Me.btnRenoncer.TabIndex = 11
        Me.btnRenoncer.Text = "Renoncer"
        Me.btnRenoncer.UseVisualStyleBackColor = True
        '
        'FrmAnnulerComptabilisation
        '
        Me.AcceptButton = Me.btnDeposer
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnRenoncer
        Me.ClientSize = New System.Drawing.Size(660, 584)
        Me.Controls.Add(Me.btnRenoncer)
        Me.Controls.Add(Me.btnDeposer)
        Me.Controls.Add(Me.lblAvertissement)
        Me.Controls.Add(Me.txtCommentaire)
        Me.Controls.Add(Me.lblCommentaire)
        Me.Controls.Add(Me.cboMotif)
        Me.Controls.Add(Me.lblMotif)
        Me.Controls.Add(Me.lblExtourne)
        Me.Controls.Add(Me.grpCoreBanking)
        Me.Controls.Add(Me.lblContenu)
        Me.Controls.Add(Me.lblJournee)
        Me.Controls.Add(Me.lblTitre)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FrmAnnulerComptabilisation"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Annuler une comptabilisation"
        Me.grpCoreBanking.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents lblTitre As System.Windows.Forms.Label
    Friend WithEvents lblJournee As System.Windows.Forms.Label
    Friend WithEvents lblContenu As System.Windows.Forms.Label
    Friend WithEvents grpCoreBanking As System.Windows.Forms.GroupBox
    Friend WithEvents lblQuestion As System.Windows.Forms.Label
    Friend WithEvents rdoNonInjecte As System.Windows.Forms.RadioButton
    Friend WithEvents rdoInjecte As System.Windows.Forms.RadioButton
    Friend WithEvents lblExtourne As System.Windows.Forms.Label
    Friend WithEvents lblMotif As System.Windows.Forms.Label
    Friend WithEvents cboMotif As System.Windows.Forms.ComboBox
    Friend WithEvents lblCommentaire As System.Windows.Forms.Label
    Friend WithEvents txtCommentaire As System.Windows.Forms.TextBox
    Friend WithEvents lblAvertissement As System.Windows.Forms.Label
    Friend WithEvents btnDeposer As System.Windows.Forms.Button
    Friend WithEvents btnRenoncer As System.Windows.Forms.Button

End Class
