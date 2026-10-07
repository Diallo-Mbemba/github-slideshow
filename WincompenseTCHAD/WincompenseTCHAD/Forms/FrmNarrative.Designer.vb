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
        Me.grpMode = New System.Windows.Forms.GroupBox()
        Me.rdoModeUnique = New System.Windows.Forms.RadioButton()
        Me.rdoModeParNature = New System.Windows.Forms.RadioButton()
        Me.lblModeIndisponible = New System.Windows.Forms.Label()
        Me.ongletsNarrative = New System.Windows.Forms.TabControl()
        Me.ongModele = New System.Windows.Forms.TabPage()
        Me.txtModele = New System.Windows.Forms.TextBox()
        Me.lblJetonsTitre = New System.Windows.Forms.Label()
        Me.lblJetons = New System.Windows.Forms.Label()
        Me.lblApercuTitre = New System.Windows.Forms.Label()
        Me.lblApercu = New System.Windows.Forms.Label()
        Me.lblLongueur = New System.Windows.Forms.Label()
        Me.lblPireCas = New System.Windows.Forms.Label()
        Me.btnDefaut = New System.Windows.Forms.Button()
        Me.ongNatures = New System.Windows.Forms.TabPage()
        Me.dgvNatures = New System.Windows.Forms.DataGridView()
        Me.lblAideNatures = New System.Windows.Forms.Label()
        Me.lblApercuNature = New System.Windows.Forms.Label()
        Me.lblLongueurNature = New System.Windows.Forms.Label()
        Me.ongJournal = New System.Windows.Forms.TabPage()
        Me.dgvJournal = New System.Windows.Forms.DataGridView()
        Me.lblJournal = New System.Windows.Forms.Label()
        Me.lblDerniereModification = New System.Windows.Forms.Label()
        Me.btnEnregistrer = New System.Windows.Forms.Button()
        Me.lblStatut = New System.Windows.Forms.Label()
        Me.btnFermer = New System.Windows.Forms.Button()
        Me.grpMode.SuspendLayout()
        Me.ongletsNarrative.SuspendLayout()
        Me.ongModele.SuspendLayout()
        Me.ongNatures.SuspendLayout()
        Me.ongJournal.SuspendLayout()
        CType(Me.dgvNatures, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvJournal, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'lblTitre
        '
        Me.lblTitre.BackColor = System.Drawing.Color.Black
        Me.lblTitre.Font = New System.Drawing.Font("Segoe UI", 14.0!, System.Drawing.FontStyle.Bold)
        Me.lblTitre.ForeColor = System.Drawing.Color.Yellow
        Me.lblTitre.Location = New System.Drawing.Point(12, 12)
        Me.lblTitre.Name = "lblTitre"
        Me.lblTitre.Size = New System.Drawing.Size(896, 34)
        Me.lblTitre.TabIndex = 0
        Me.lblTitre.Text = "   Narrative comptable"
        Me.lblTitre.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblSousTitre
        '
        Me.lblSousTitre.ForeColor = System.Drawing.SystemColors.GrayText
        Me.lblSousTitre.Location = New System.Drawing.Point(12, 50)
        Me.lblSousTitre.Name = "lblSousTitre"
        Me.lblSousTitre.Size = New System.Drawing.Size(896, 34)
        Me.lblSousTitre.TabIndex = 1
        Me.lblSousTitre.Text = "Deux textes : les LIBELLÉS DES LIGNES de la pièce comptable, et le MODÈLE DU POIN" &
            "T DE VENTE que le core banking reçoit dans ADDLTEXT. Ils valent pour tous les pos" &
            "tes et pour toutes les journées à venir ; les pièces déjà produites gardent les l" &
            "eurs."
        '
        'grpMode
        '
        Me.grpMode.Controls.Add(Me.rdoModeUnique)
        Me.grpMode.Controls.Add(Me.rdoModeParNature)
        Me.grpMode.Controls.Add(Me.lblModeIndisponible)
        Me.grpMode.Location = New System.Drawing.Point(12, 90)
        Me.grpMode.Name = "grpMode"
        Me.grpMode.Size = New System.Drawing.Size(896, 100)
        Me.grpMode.TabIndex = 2
        Me.grpMode.TabStop = False
        Me.grpMode.Text = "Comment les libellés de la PIÈCE sont choisis (le fichier core banking, lui, port" &
            "e toujours le modèle du point de vente)"
        '
        'rdoModeUnique
        '
        Me.rdoModeUnique.Location = New System.Drawing.Point(18, 24)
        Me.rdoModeUnique.Name = "rdoModeUnique"
        Me.rdoModeUnique.Size = New System.Drawing.Size(860, 22)
        Me.rdoModeUnique.TabIndex = 0
        Me.rdoModeUnique.Text = "UN SEUL LIBELLÉ pour les douze lignes d'un point de vente — celui du modèle du po" &
            "int de vente, ci-dessous."
        Me.rdoModeUnique.UseVisualStyleBackColor = True
        '
        'rdoModeParNature
        '
        Me.rdoModeParNature.Checked = True
        Me.rdoModeParNature.Location = New System.Drawing.Point(18, 48)
        Me.rdoModeParNature.Name = "rdoModeParNature"
        Me.rdoModeParNature.Size = New System.Drawing.Size(860, 22)
        Me.rdoModeParNature.TabIndex = 1
        Me.rdoModeParNature.TabStop = True
        Me.rdoModeParNature.Text = "UN LIBELLÉ PAR NATURE de mouvement — mouvement, contrepartie, commissions, taxes," &
            " écart d'arrondi. C'est la forme habituelle de la pièce de la banque."
        Me.rdoModeParNature.UseVisualStyleBackColor = True
        '
        'lblModeIndisponible
        '
        Me.lblModeIndisponible.ForeColor = System.Drawing.Color.Firebrick
        Me.lblModeIndisponible.Location = New System.Drawing.Point(36, 72)
        Me.lblModeIndisponible.Name = "lblModeIndisponible"
        Me.lblModeIndisponible.Size = New System.Drawing.Size(842, 22)
        Me.lblModeIndisponible.TabIndex = 2
        Me.lblModeIndisponible.Visible = False
        '
        'ongletsNarrative
        '
        Me.ongletsNarrative.Controls.Add(Me.ongModele)
        Me.ongletsNarrative.Controls.Add(Me.ongNatures)
        Me.ongletsNarrative.Controls.Add(Me.ongJournal)
        Me.ongletsNarrative.Location = New System.Drawing.Point(12, 198)
        Me.ongletsNarrative.Name = "ongletsNarrative"
        Me.ongletsNarrative.SelectedIndex = 0
        Me.ongletsNarrative.Size = New System.Drawing.Size(896, 380)
        Me.ongletsNarrative.TabIndex = 3
        '
        'ongModele
        '
        Me.ongModele.Controls.Add(Me.txtModele)
        Me.ongModele.Controls.Add(Me.lblJetonsTitre)
        Me.ongModele.Controls.Add(Me.lblJetons)
        Me.ongModele.Controls.Add(Me.lblApercuTitre)
        Me.ongModele.Controls.Add(Me.lblApercu)
        Me.ongModele.Controls.Add(Me.lblLongueur)
        Me.ongModele.Controls.Add(Me.lblPireCas)
        Me.ongModele.Controls.Add(Me.btnDefaut)
        Me.ongModele.Location = New System.Drawing.Point(4, 22)
        Me.ongModele.Name = "ongModele"
        Me.ongModele.Padding = New System.Windows.Forms.Padding(3)
        Me.ongModele.Size = New System.Drawing.Size(888, 354)
        Me.ongModele.TabIndex = 0
        Me.ongModele.Text = "Le modèle du point de vente"
        Me.ongModele.UseVisualStyleBackColor = True
        '
        'txtModele
        '
        Me.txtModele.Location = New System.Drawing.Point(18, 20)
        Me.txtModele.Name = "txtModele"
        Me.txtModele.Size = New System.Drawing.Size(850, 20)
        Me.txtModele.TabIndex = 0
        '
        'lblJetonsTitre
        '
        Me.lblJetonsTitre.Location = New System.Drawing.Point(18, 52)
        Me.lblJetonsTitre.Name = "lblJetonsTitre"
        Me.lblJetonsTitre.Size = New System.Drawing.Size(850, 18)
        Me.lblJetonsTitre.TabIndex = 1
        Me.lblJetonsTitre.Text = "Repères reconnus — tout le reste est écrit tel quel :"
        '
        'lblJetons
        '
        Me.lblJetons.ForeColor = System.Drawing.SystemColors.GrayText
        Me.lblJetons.Location = New System.Drawing.Point(30, 74)
        Me.lblJetons.Name = "lblJetons"
        Me.lblJetons.Size = New System.Drawing.Size(838, 58)
        Me.lblJetons.TabIndex = 2
        '
        'lblApercuTitre
        '
        Me.lblApercuTitre.Location = New System.Drawing.Point(18, 140)
        Me.lblApercuTitre.Name = "lblApercuTitre"
        Me.lblApercuTitre.Size = New System.Drawing.Size(850, 18)
        Me.lblApercuTitre.TabIndex = 3
        Me.lblApercuTitre.Text = "Aperçu — la phrase telle qu'elle sortira, dans le cas le plus long :"
        '
        'lblApercu
        '
        Me.lblApercu.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold)
        Me.lblApercu.Location = New System.Drawing.Point(30, 162)
        Me.lblApercu.Name = "lblApercu"
        Me.lblApercu.Size = New System.Drawing.Size(838, 40)
        Me.lblApercu.TabIndex = 4
        '
        'lblLongueur
        '
        Me.lblLongueur.Location = New System.Drawing.Point(18, 208)
        Me.lblLongueur.Name = "lblLongueur"
        Me.lblLongueur.Size = New System.Drawing.Size(850, 20)
        Me.lblLongueur.TabIndex = 5
        '
        'lblPireCas
        '
        Me.lblPireCas.ForeColor = System.Drawing.SystemColors.GrayText
        Me.lblPireCas.Location = New System.Drawing.Point(18, 232)
        Me.lblPireCas.Name = "lblPireCas"
        Me.lblPireCas.Size = New System.Drawing.Size(850, 46)
        Me.lblPireCas.TabIndex = 6
        '
        'btnDefaut
        '
        Me.btnDefaut.Location = New System.Drawing.Point(18, 288)
        Me.btnDefaut.Name = "btnDefaut"
        Me.btnDefaut.Size = New System.Drawing.Size(230, 32)
        Me.btnDefaut.TabIndex = 7
        Me.btnDefaut.Text = "Rétablir le modèle par défaut"
        Me.btnDefaut.UseVisualStyleBackColor = True
        '
        'ongNatures
        '
        Me.ongNatures.Controls.Add(Me.dgvNatures)
        Me.ongNatures.Controls.Add(Me.lblApercuNature)
        Me.ongNatures.Controls.Add(Me.lblLongueurNature)
        Me.ongNatures.Controls.Add(Me.lblAideNatures)
        Me.ongNatures.Location = New System.Drawing.Point(4, 22)
        Me.ongNatures.Name = "ongNatures"
        Me.ongNatures.Padding = New System.Windows.Forms.Padding(3)
        Me.ongNatures.Size = New System.Drawing.Size(888, 354)
        Me.ongNatures.TabIndex = 1
        Me.ongNatures.Text = "Les libellés par nature"
        Me.ongNatures.UseVisualStyleBackColor = True
        '
        'dgvNatures
        '
        Me.dgvNatures.AllowUserToAddRows = False
        Me.dgvNatures.AllowUserToDeleteRows = False
        Me.dgvNatures.AllowUserToResizeRows = False
        Me.dgvNatures.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.None
        Me.dgvNatures.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvNatures.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter
        Me.dgvNatures.Location = New System.Drawing.Point(12, 12)
        Me.dgvNatures.MultiSelect = False
        Me.dgvNatures.Name = "dgvNatures"
        Me.dgvNatures.RowHeadersVisible = False
        Me.dgvNatures.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect
        Me.dgvNatures.Size = New System.Drawing.Size(864, 240)
        Me.dgvNatures.TabIndex = 0
        '
        'lblApercuNature
        '
        Me.lblApercuNature.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold)
        Me.lblApercuNature.Location = New System.Drawing.Point(12, 258)
        Me.lblApercuNature.Name = "lblApercuNature"
        Me.lblApercuNature.Size = New System.Drawing.Size(864, 34)
        Me.lblApercuNature.TabIndex = 1
        '
        'lblLongueurNature
        '
        Me.lblLongueurNature.Location = New System.Drawing.Point(12, 294)
        Me.lblLongueurNature.Name = "lblLongueurNature"
        Me.lblLongueurNature.Size = New System.Drawing.Size(864, 20)
        Me.lblLongueurNature.TabIndex = 2
        '
        'lblAideNatures
        '
        Me.lblAideNatures.ForeColor = System.Drawing.SystemColors.GrayText
        Me.lblAideNatures.Location = New System.Drawing.Point(12, 316)
        Me.lblAideNatures.Name = "lblAideNatures"
        Me.lblAideNatures.Size = New System.Drawing.Size(864, 34)
        Me.lblAideNatures.TabIndex = 3
        '
        'ongJournal
        '
        Me.ongJournal.Controls.Add(Me.dgvJournal)
        Me.ongJournal.Controls.Add(Me.lblJournal)
        Me.ongJournal.Location = New System.Drawing.Point(4, 22)
        Me.ongJournal.Name = "ongJournal"
        Me.ongJournal.Padding = New System.Windows.Forms.Padding(3)
        Me.ongJournal.Size = New System.Drawing.Size(888, 354)
        Me.ongJournal.TabIndex = 2
        Me.ongJournal.Text = "Historique des modifications"
        Me.ongJournal.UseVisualStyleBackColor = True
        '
        'dgvJournal
        '
        Me.dgvJournal.AllowUserToAddRows = False
        Me.dgvJournal.AllowUserToDeleteRows = False
        Me.dgvJournal.AllowUserToResizeRows = False
        Me.dgvJournal.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.None
        Me.dgvJournal.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvJournal.Location = New System.Drawing.Point(12, 12)
        Me.dgvJournal.MultiSelect = False
        Me.dgvJournal.Name = "dgvJournal"
        Me.dgvJournal.ReadOnly = True
        Me.dgvJournal.RowHeadersVisible = False
        Me.dgvJournal.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvJournal.Size = New System.Drawing.Size(864, 290)
        Me.dgvJournal.TabIndex = 0
        '
        'lblJournal
        '
        Me.lblJournal.ForeColor = System.Drawing.SystemColors.GrayText
        Me.lblJournal.Location = New System.Drawing.Point(12, 310)
        Me.lblJournal.Name = "lblJournal"
        Me.lblJournal.Size = New System.Drawing.Size(864, 40)
        Me.lblJournal.TabIndex = 1
        '
        'lblDerniereModification
        '
        Me.lblDerniereModification.ForeColor = System.Drawing.SystemColors.GrayText
        Me.lblDerniereModification.Location = New System.Drawing.Point(12, 586)
        Me.lblDerniereModification.Name = "lblDerniereModification"
        Me.lblDerniereModification.Size = New System.Drawing.Size(896, 20)
        Me.lblDerniereModification.TabIndex = 4
        '
        'btnEnregistrer
        '
        Me.btnEnregistrer.Location = New System.Drawing.Point(12, 612)
        Me.btnEnregistrer.Name = "btnEnregistrer"
        Me.btnEnregistrer.Size = New System.Drawing.Size(180, 32)
        Me.btnEnregistrer.TabIndex = 5
        Me.btnEnregistrer.Text = "Enregistrer"
        Me.btnEnregistrer.UseVisualStyleBackColor = True
        '
        'lblStatut
        '
        Me.lblStatut.ForeColor = System.Drawing.SystemColors.GrayText
        Me.lblStatut.Location = New System.Drawing.Point(202, 618)
        Me.lblStatut.Name = "lblStatut"
        Me.lblStatut.Size = New System.Drawing.Size(600, 20)
        Me.lblStatut.TabIndex = 6
        '
        'btnFermer
        '
        Me.btnFermer.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnFermer.Location = New System.Drawing.Point(816, 612)
        Me.btnFermer.Name = "btnFermer"
        Me.btnFermer.Size = New System.Drawing.Size(92, 32)
        Me.btnFermer.TabIndex = 7
        Me.btnFermer.Text = "Fermer"
        Me.btnFermer.UseVisualStyleBackColor = True
        '
        'FrmNarrative
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnFermer
        Me.ClientSize = New System.Drawing.Size(920, 658)
        Me.Controls.Add(Me.btnFermer)
        Me.Controls.Add(Me.lblStatut)
        Me.Controls.Add(Me.btnEnregistrer)
        Me.Controls.Add(Me.lblDerniereModification)
        Me.Controls.Add(Me.ongletsNarrative)
        Me.Controls.Add(Me.grpMode)
        Me.Controls.Add(Me.lblSousTitre)
        Me.Controls.Add(Me.lblTitre)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FrmNarrative"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Narrative comptable"
        Me.grpMode.ResumeLayout(False)
        Me.ongModele.ResumeLayout(False)
        Me.ongNatures.ResumeLayout(False)
        Me.ongJournal.ResumeLayout(False)
        Me.ongletsNarrative.ResumeLayout(False)
        CType(Me.dgvNatures, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvJournal, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents lblTitre As System.Windows.Forms.Label
    Friend WithEvents lblSousTitre As System.Windows.Forms.Label
    Friend WithEvents grpMode As System.Windows.Forms.GroupBox
    Friend WithEvents rdoModeUnique As System.Windows.Forms.RadioButton
    Friend WithEvents rdoModeParNature As System.Windows.Forms.RadioButton
    Friend WithEvents lblModeIndisponible As System.Windows.Forms.Label
    Friend WithEvents ongletsNarrative As System.Windows.Forms.TabControl
    Friend WithEvents ongModele As System.Windows.Forms.TabPage
    Friend WithEvents txtModele As System.Windows.Forms.TextBox
    Friend WithEvents lblJetonsTitre As System.Windows.Forms.Label
    Friend WithEvents lblJetons As System.Windows.Forms.Label
    Friend WithEvents lblApercuTitre As System.Windows.Forms.Label
    Friend WithEvents lblApercu As System.Windows.Forms.Label
    Friend WithEvents lblLongueur As System.Windows.Forms.Label
    Friend WithEvents lblPireCas As System.Windows.Forms.Label
    Friend WithEvents btnDefaut As System.Windows.Forms.Button
    Friend WithEvents ongNatures As System.Windows.Forms.TabPage
    Friend WithEvents dgvNatures As System.Windows.Forms.DataGridView
    Friend WithEvents lblApercuNature As System.Windows.Forms.Label
    Friend WithEvents lblLongueurNature As System.Windows.Forms.Label
    Friend WithEvents lblAideNatures As System.Windows.Forms.Label
    Friend WithEvents ongJournal As System.Windows.Forms.TabPage
    Friend WithEvents dgvJournal As System.Windows.Forms.DataGridView
    Friend WithEvents lblJournal As System.Windows.Forms.Label
    Friend WithEvents lblDerniereModification As System.Windows.Forms.Label
    Friend WithEvents btnEnregistrer As System.Windows.Forms.Button
    Friend WithEvents lblStatut As System.Windows.Forms.Label
    Friend WithEvents btnFermer As System.Windows.Forms.Button

End Class
