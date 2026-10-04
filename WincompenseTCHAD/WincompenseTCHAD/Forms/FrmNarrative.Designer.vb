Option Strict On
Option Explicit On

Partial Class FrmNarrative
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
        Me.grpModele = New System.Windows.Forms.GroupBox()
        Me.txtModele = New System.Windows.Forms.TextBox()
        Me.lblJetonsTitre = New System.Windows.Forms.Label()
        Me.lblJetons = New System.Windows.Forms.Label()
        Me.grpApercu = New System.Windows.Forms.GroupBox()
        Me.lblApercu = New System.Windows.Forms.Label()
        Me.lblLongueur = New System.Windows.Forms.Label()
        Me.lblPireCas = New System.Windows.Forms.Label()
        Me.lblDerniereModification = New System.Windows.Forms.Label()
        Me.btnDefaut = New System.Windows.Forms.Button()
        Me.btnEnregistrer = New System.Windows.Forms.Button()
        Me.lblStatut = New System.Windows.Forms.Label()
        Me.btnFermer = New System.Windows.Forms.Button()
        Me.grpModele.SuspendLayout()
        Me.grpApercu.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblTitre
        '
        Me.lblTitre.BackColor = System.Drawing.Color.Black
        Me.lblTitre.Font = New System.Drawing.Font("Segoe UI", 14.0!, System.Drawing.FontStyle.Bold)
        Me.lblTitre.ForeColor = System.Drawing.Color.Yellow
        Me.lblTitre.Location = New System.Drawing.Point(12, 12)
        Me.lblTitre.Name = "lblTitre"
        Me.lblTitre.Size = New System.Drawing.Size(736, 34)
        Me.lblTitre.TabIndex = 0
        Me.lblTitre.Text = "   Narrative comptable"
        Me.lblTitre.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblSousTitre
        '
        Me.lblSousTitre.ForeColor = System.Drawing.SystemColors.GrayText
        Me.lblSousTitre.Location = New System.Drawing.Point(12, 50)
        Me.lblSousTitre.Name = "lblSousTitre"
        Me.lblSousTitre.Size = New System.Drawing.Size(736, 46)
        Me.lblSousTitre.TabIndex = 1
        Me.lblSousTitre.Text = "Ce texte est le libellé de CHAQUE ligne des pièces comptables, et il part tel quel" &
            " dans la colonne ADDLTEXT du fichier chargé au core banking. Il vaut pour tous le" &
            "s postes et pour toutes les journées à venir ; les pièces déjà produites gardent " &
            "le leur."
        '
        'grpModele
        '
        Me.grpModele.Controls.Add(Me.txtModele)
        Me.grpModele.Controls.Add(Me.lblJetonsTitre)
        Me.grpModele.Controls.Add(Me.lblJetons)
        Me.grpModele.Location = New System.Drawing.Point(12, 102)
        Me.grpModele.Name = "grpModele"
        Me.grpModele.Size = New System.Drawing.Size(736, 150)
        Me.grpModele.TabIndex = 2
        Me.grpModele.TabStop = False
        Me.grpModele.Text = "Le modèle"
        '
        'txtModele
        '
        Me.txtModele.Location = New System.Drawing.Point(18, 32)
        Me.txtModele.Name = "txtModele"
        Me.txtModele.Size = New System.Drawing.Size(700, 20)
        Me.txtModele.TabIndex = 0
        '
        'lblJetonsTitre
        '
        Me.lblJetonsTitre.Location = New System.Drawing.Point(18, 64)
        Me.lblJetonsTitre.Name = "lblJetonsTitre"
        Me.lblJetonsTitre.Size = New System.Drawing.Size(700, 18)
        Me.lblJetonsTitre.TabIndex = 1
        Me.lblJetonsTitre.Text = "Repères reconnus — tout le reste est écrit tel quel :"
        '
        'lblJetons
        '
        Me.lblJetons.ForeColor = System.Drawing.SystemColors.GrayText
        Me.lblJetons.Location = New System.Drawing.Point(30, 86)
        Me.lblJetons.Name = "lblJetons"
        Me.lblJetons.Size = New System.Drawing.Size(688, 58)
        Me.lblJetons.TabIndex = 2
        '
        'grpApercu
        '
        Me.grpApercu.Controls.Add(Me.lblApercu)
        Me.grpApercu.Controls.Add(Me.lblLongueur)
        Me.grpApercu.Controls.Add(Me.lblPireCas)
        Me.grpApercu.Location = New System.Drawing.Point(12, 262)
        Me.grpApercu.Name = "grpApercu"
        Me.grpApercu.Size = New System.Drawing.Size(736, 152)
        Me.grpApercu.TabIndex = 3
        Me.grpApercu.TabStop = False
        Me.grpApercu.Text = "Aperçu — la phrase telle qu'elle sortira, dans le cas le plus long"
        '
        'lblApercu
        '
        Me.lblApercu.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold)
        Me.lblApercu.Location = New System.Drawing.Point(18, 28)
        Me.lblApercu.Name = "lblApercu"
        Me.lblApercu.Size = New System.Drawing.Size(700, 40)
        Me.lblApercu.TabIndex = 0
        '
        'lblLongueur
        '
        Me.lblLongueur.Location = New System.Drawing.Point(18, 74)
        Me.lblLongueur.Name = "lblLongueur"
        Me.lblLongueur.Size = New System.Drawing.Size(700, 20)
        Me.lblLongueur.TabIndex = 1
        '
        'lblPireCas
        '
        Me.lblPireCas.ForeColor = System.Drawing.SystemColors.GrayText
        Me.lblPireCas.Location = New System.Drawing.Point(18, 98)
        Me.lblPireCas.Name = "lblPireCas"
        Me.lblPireCas.Size = New System.Drawing.Size(700, 46)
        Me.lblPireCas.TabIndex = 2
        '
        'lblDerniereModification
        '
        Me.lblDerniereModification.ForeColor = System.Drawing.SystemColors.GrayText
        Me.lblDerniereModification.Location = New System.Drawing.Point(12, 422)
        Me.lblDerniereModification.Name = "lblDerniereModification"
        Me.lblDerniereModification.Size = New System.Drawing.Size(736, 34)
        Me.lblDerniereModification.TabIndex = 4
        '
        'btnDefaut
        '
        Me.btnDefaut.Location = New System.Drawing.Point(12, 466)
        Me.btnDefaut.Name = "btnDefaut"
        Me.btnDefaut.Size = New System.Drawing.Size(230, 32)
        Me.btnDefaut.TabIndex = 5
        Me.btnDefaut.Text = "Rétablir le modèle par défaut"
        Me.btnDefaut.UseVisualStyleBackColor = True
        '
        'btnEnregistrer
        '
        Me.btnEnregistrer.Location = New System.Drawing.Point(252, 466)
        Me.btnEnregistrer.Name = "btnEnregistrer"
        Me.btnEnregistrer.Size = New System.Drawing.Size(150, 32)
        Me.btnEnregistrer.TabIndex = 6
        Me.btnEnregistrer.Text = "Enregistrer"
        Me.btnEnregistrer.UseVisualStyleBackColor = True
        '
        'lblStatut
        '
        Me.lblStatut.ForeColor = System.Drawing.SystemColors.GrayText
        Me.lblStatut.Location = New System.Drawing.Point(412, 472)
        Me.lblStatut.Name = "lblStatut"
        Me.lblStatut.Size = New System.Drawing.Size(236, 20)
        Me.lblStatut.TabIndex = 7
        '
        'btnFermer
        '
        Me.btnFermer.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnFermer.Location = New System.Drawing.Point(656, 466)
        Me.btnFermer.Name = "btnFermer"
        Me.btnFermer.Size = New System.Drawing.Size(92, 32)
        Me.btnFermer.TabIndex = 8
        Me.btnFermer.Text = "Fermer"
        Me.btnFermer.UseVisualStyleBackColor = True
        '
        'FrmNarrative
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnFermer
        Me.ClientSize = New System.Drawing.Size(760, 512)
        Me.Controls.Add(Me.btnFermer)
        Me.Controls.Add(Me.lblStatut)
        Me.Controls.Add(Me.btnEnregistrer)
        Me.Controls.Add(Me.btnDefaut)
        Me.Controls.Add(Me.lblDerniereModification)
        Me.Controls.Add(Me.grpApercu)
        Me.Controls.Add(Me.grpModele)
        Me.Controls.Add(Me.lblSousTitre)
        Me.Controls.Add(Me.lblTitre)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FrmNarrative"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Narrative comptable"
        Me.grpModele.ResumeLayout(False)
        Me.grpApercu.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents lblTitre As System.Windows.Forms.Label
    Friend WithEvents lblSousTitre As System.Windows.Forms.Label
    Friend WithEvents grpModele As System.Windows.Forms.GroupBox
    Friend WithEvents txtModele As System.Windows.Forms.TextBox
    Friend WithEvents lblJetonsTitre As System.Windows.Forms.Label
    Friend WithEvents lblJetons As System.Windows.Forms.Label
    Friend WithEvents grpApercu As System.Windows.Forms.GroupBox
    Friend WithEvents lblApercu As System.Windows.Forms.Label
    Friend WithEvents lblLongueur As System.Windows.Forms.Label
    Friend WithEvents lblPireCas As System.Windows.Forms.Label
    Friend WithEvents lblDerniereModification As System.Windows.Forms.Label
    Friend WithEvents btnDefaut As System.Windows.Forms.Button
    Friend WithEvents btnEnregistrer As System.Windows.Forms.Button
    Friend WithEvents lblStatut As System.Windows.Forms.Label
    Friend WithEvents btnFermer As System.Windows.Forms.Button

End Class
