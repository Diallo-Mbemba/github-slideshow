Option Strict On
Option Explicit On

Partial Class FrmOptionsTraitement
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
        Me.grpVisa = New System.Windows.Forms.GroupBox()
        Me.rdoVisaOui = New System.Windows.Forms.RadioButton()
        Me.rdoVisaNon = New System.Windows.Forms.RadioButton()
        Me.lblConsequence = New System.Windows.Forms.Label()
        Me.lblDerniereModification = New System.Windows.Forms.Label()
        Me.btnEnregistrer = New System.Windows.Forms.Button()
        Me.lblStatut = New System.Windows.Forms.Label()
        Me.btnFermer = New System.Windows.Forms.Button()
        Me.grpVisa.SuspendLayout()
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
        Me.lblTitre.Text = "   Options de traitement"
        Me.lblTitre.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblSousTitre
        '
        Me.lblSousTitre.ForeColor = System.Drawing.SystemColors.GrayText
        Me.lblSousTitre.Location = New System.Drawing.Point(12, 50)
        Me.lblSousTitre.Name = "lblSousTitre"
        Me.lblSousTitre.Size = New System.Drawing.Size(636, 34)
        Me.lblSousTitre.TabIndex = 1
        Me.lblSousTitre.Text = "Ces options décrivent la façon de travailler de la banque, et non la configuration" &
            " d'un poste : elles valent pour tout le monde et sont conservées dans la base."
        '
        'grpVisa
        '
        Me.grpVisa.Controls.Add(Me.rdoVisaOui)
        Me.grpVisa.Controls.Add(Me.rdoVisaNon)
        Me.grpVisa.Controls.Add(Me.lblConsequence)
        Me.grpVisa.Location = New System.Drawing.Point(12, 92)
        Me.grpVisa.Name = "grpVisa"
        Me.grpVisa.Size = New System.Drawing.Size(636, 176)
        Me.grpVisa.TabIndex = 2
        Me.grpVisa.TabStop = False
        Me.grpVisa.Text = "Le visa est-il obligatoire avant de produire le fichier core banking ?"
        '
        'rdoVisaNon
        '
        Me.rdoVisaNon.Checked = True
        Me.rdoVisaNon.Location = New System.Drawing.Point(18, 28)
        Me.rdoVisaNon.Name = "rdoVisaNon"
        Me.rdoVisaNon.Size = New System.Drawing.Size(600, 40)
        Me.rdoVisaNon.TabIndex = 0
        Me.rdoVisaNon.TabStop = True
        Me.rdoVisaNon.Text = "NON — l'application AVERTIT que la journée n'est pas visée, et laisse l'agent déci" &
            "der de produire le fichier."
        Me.rdoVisaNon.UseVisualStyleBackColor = True
        '
        'rdoVisaOui
        '
        Me.rdoVisaOui.Location = New System.Drawing.Point(18, 72)
        Me.rdoVisaOui.Name = "rdoVisaOui"
        Me.rdoVisaOui.Size = New System.Drawing.Size(600, 40)
        Me.rdoVisaOui.TabIndex = 1
        Me.rdoVisaOui.Text = "OUI — l'application REFUSE de produire le fichier tant qu'un supérieur n'a pas vis" &
            "é la journée."
        Me.rdoVisaOui.UseVisualStyleBackColor = True
        '
        'lblConsequence
        '
        Me.lblConsequence.Location = New System.Drawing.Point(18, 118)
        Me.lblConsequence.Name = "lblConsequence"
        Me.lblConsequence.Size = New System.Drawing.Size(600, 48)
        Me.lblConsequence.TabIndex = 2
        '
        'lblDerniereModification
        '
        Me.lblDerniereModification.ForeColor = System.Drawing.SystemColors.GrayText
        Me.lblDerniereModification.Location = New System.Drawing.Point(12, 276)
        Me.lblDerniereModification.Name = "lblDerniereModification"
        Me.lblDerniereModification.Size = New System.Drawing.Size(636, 20)
        Me.lblDerniereModification.TabIndex = 3
        '
        'btnEnregistrer
        '
        Me.btnEnregistrer.Location = New System.Drawing.Point(12, 308)
        Me.btnEnregistrer.Name = "btnEnregistrer"
        Me.btnEnregistrer.Size = New System.Drawing.Size(160, 32)
        Me.btnEnregistrer.TabIndex = 4
        Me.btnEnregistrer.Text = "Enregistrer"
        Me.btnEnregistrer.UseVisualStyleBackColor = True
        '
        'lblStatut
        '
        Me.lblStatut.ForeColor = System.Drawing.SystemColors.GrayText
        Me.lblStatut.Location = New System.Drawing.Point(182, 314)
        Me.lblStatut.Name = "lblStatut"
        Me.lblStatut.Size = New System.Drawing.Size(370, 20)
        Me.lblStatut.TabIndex = 5
        '
        'btnFermer
        '
        Me.btnFermer.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnFermer.Location = New System.Drawing.Point(556, 308)
        Me.btnFermer.Name = "btnFermer"
        Me.btnFermer.Size = New System.Drawing.Size(92, 32)
        Me.btnFermer.TabIndex = 6
        Me.btnFermer.Text = "Fermer"
        Me.btnFermer.UseVisualStyleBackColor = True
        '
        'FrmOptionsTraitement
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnFermer
        Me.ClientSize = New System.Drawing.Size(660, 352)
        Me.Controls.Add(Me.btnFermer)
        Me.Controls.Add(Me.lblStatut)
        Me.Controls.Add(Me.btnEnregistrer)
        Me.Controls.Add(Me.lblDerniereModification)
        Me.Controls.Add(Me.grpVisa)
        Me.Controls.Add(Me.lblSousTitre)
        Me.Controls.Add(Me.lblTitre)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FrmOptionsTraitement"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Options de traitement"
        Me.grpVisa.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents lblTitre As System.Windows.Forms.Label
    Friend WithEvents lblSousTitre As System.Windows.Forms.Label
    Friend WithEvents grpVisa As System.Windows.Forms.GroupBox
    Friend WithEvents rdoVisaOui As System.Windows.Forms.RadioButton
    Friend WithEvents rdoVisaNon As System.Windows.Forms.RadioButton
    Friend WithEvents lblConsequence As System.Windows.Forms.Label
    Friend WithEvents lblDerniereModification As System.Windows.Forms.Label
    Friend WithEvents btnEnregistrer As System.Windows.Forms.Button
    Friend WithEvents lblStatut As System.Windows.Forms.Label
    Friend WithEvents btnFermer As System.Windows.Forms.Button

End Class
