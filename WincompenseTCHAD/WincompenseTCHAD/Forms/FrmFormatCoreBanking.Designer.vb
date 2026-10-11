Option Strict On
Option Explicit On

Partial Class FrmFormatCoreBanking
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
        Me.ongletsFormat = New System.Windows.Forms.TabControl()
        Me.ongCodes = New System.Windows.Forms.TabPage()
        Me.lblDetbsjrnl = New System.Windows.Forms.Label()
        Me.txtDetbsjrnl = New System.Windows.Forms.TextBox()
        Me.lblBrn = New System.Windows.Forms.Label()
        Me.txtBrn = New System.Windows.Forms.TextBox()
        Me.lblSrccode = New System.Windows.Forms.Label()
        Me.txtSrccode = New System.Windows.Forms.TextBox()
        Me.lblTxncdDebit = New System.Windows.Forms.Label()
        Me.txtTxncdDebit = New System.Windows.Forms.TextBox()
        Me.lblTxncdCredit = New System.Windows.Forms.Label()
        Me.txtTxncdCredit = New System.Windows.Forms.TextBox()
        Me.lblAgenceSiege = New System.Windows.Forms.Label()
        Me.txtAgenceSiege = New System.Windows.Forms.TextBox()
        Me.lblAvertissementBrn = New System.Windows.Forms.Label()
        Me.btnDefaut = New System.Windows.Forms.Button()
        Me.ongComptes = New System.Windows.Forms.TabPage()
        Me.lblAideComptes = New System.Windows.Forms.Label()
        Me.dgvComptes = New System.Windows.Forms.DataGridView()
        Me.ongJournal = New System.Windows.Forms.TabPage()
        Me.lblJournal = New System.Windows.Forms.Label()
        Me.dgvJournal = New System.Windows.Forms.DataGridView()
        Me.lblApercuTitre = New System.Windows.Forms.Label()
        Me.lblApercu = New System.Windows.Forms.Label()
        Me.lblDerniereModification = New System.Windows.Forms.Label()
        Me.lblStatut = New System.Windows.Forms.Label()
        Me.btnEnregistrer = New System.Windows.Forms.Button()
        Me.btnFermer = New System.Windows.Forms.Button()
        Me.ongletsFormat.SuspendLayout()
        Me.ongCodes.SuspendLayout()
        Me.ongComptes.SuspendLayout()
        Me.ongJournal.SuspendLayout()
        CType(Me.dgvComptes, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvJournal, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'lblTitre
        '
        Me.lblTitre.BackColor = System.Drawing.Color.Black
        Me.lblTitre.Font = New System.Drawing.Font("Segoe UI", 14.25!, System.Drawing.FontStyle.Bold)
        Me.lblTitre.ForeColor = System.Drawing.Color.Gold
        Me.lblTitre.Location = New System.Drawing.Point(12, 12)
        Me.lblTitre.Name = "lblTitre"
        Me.lblTitre.Size = New System.Drawing.Size(896, 34)
        Me.lblTitre.TabIndex = 0
        Me.lblTitre.Text = "   Format du fichier core banking"
        Me.lblTitre.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblSousTitre
        '
        Me.lblSousTitre.ForeColor = System.Drawing.SystemColors.GrayText
        Me.lblSousTitre.Location = New System.Drawing.Point(12, 48)
        Me.lblSousTitre.Name = "lblSousTitre"
        Me.lblSousTitre.Size = New System.Drawing.Size(896, 40)
        Me.lblSousTitre.TabIndex = 1
        Me.lblSousTitre.Text = "Les codes que la banque a fixés pour son fichier d'interface. Ils partent sur TOU" &
            "TES les lignes de tous les lots à venir ; les fichiers déjà chargés portent ceux " &
            "qui étaient en vigueur ce jour-là, et ne changent pas."
        '
        'ongletsFormat
        '
        Me.ongletsFormat.Controls.Add(Me.ongCodes)
        Me.ongletsFormat.Controls.Add(Me.ongComptes)
        Me.ongletsFormat.Controls.Add(Me.ongJournal)
        Me.ongletsFormat.Location = New System.Drawing.Point(12, 92)
        Me.ongletsFormat.Name = "ongletsFormat"
        Me.ongletsFormat.SelectedIndex = 0
        Me.ongletsFormat.Size = New System.Drawing.Size(896, 440)
        Me.ongletsFormat.TabIndex = 2
        '
        'ongCodes
        '
        Me.ongCodes.Controls.Add(Me.lblDetbsjrnl)
        Me.ongCodes.Controls.Add(Me.txtDetbsjrnl)
        Me.ongCodes.Controls.Add(Me.lblBrn)
        Me.ongCodes.Controls.Add(Me.txtBrn)
        Me.ongCodes.Controls.Add(Me.lblSrccode)
        Me.ongCodes.Controls.Add(Me.txtSrccode)
        Me.ongCodes.Controls.Add(Me.lblTxncdDebit)
        Me.ongCodes.Controls.Add(Me.txtTxncdDebit)
        Me.ongCodes.Controls.Add(Me.lblTxncdCredit)
        Me.ongCodes.Controls.Add(Me.txtTxncdCredit)
        Me.ongCodes.Controls.Add(Me.lblAgenceSiege)
        Me.ongCodes.Controls.Add(Me.txtAgenceSiege)
        Me.ongCodes.Controls.Add(Me.lblAvertissementBrn)
        Me.ongCodes.Controls.Add(Me.btnDefaut)
        Me.ongCodes.Location = New System.Drawing.Point(4, 22)
        Me.ongCodes.Name = "ongCodes"
        Me.ongCodes.Padding = New System.Windows.Forms.Padding(3)
        Me.ongCodes.Size = New System.Drawing.Size(888, 414)
        Me.ongCodes.TabIndex = 0
        Me.ongCodes.Text = "Les codes du fichier"
        Me.ongCodes.UseVisualStyleBackColor = True
        '
        'lblDetbsjrnl
        '
        Me.lblDetbsjrnl.Location = New System.Drawing.Point(18, 24)
        Me.lblDetbsjrnl.Name = "lblDetbsjrnl"
        Me.lblDetbsjrnl.Size = New System.Drawing.Size(300, 20)
        Me.lblDetbsjrnl.TabIndex = 0
        Me.lblDetbsjrnl.Text = "DETBSJRNL — journal, sur toutes les lignes :"
        '
        'txtDetbsjrnl
        '
        Me.txtDetbsjrnl.Location = New System.Drawing.Point(324, 21)
        Me.txtDetbsjrnl.Name = "txtDetbsjrnl"
        Me.txtDetbsjrnl.Size = New System.Drawing.Size(160, 20)
        Me.txtDetbsjrnl.TabIndex = 1
        '
        'lblBrn
        '
        Me.lblBrn.Location = New System.Drawing.Point(18, 58)
        Me.lblBrn.Name = "lblBrn"
        Me.lblBrn.Size = New System.Drawing.Size(300, 20)
        Me.lblBrn.TabIndex = 2
        Me.lblBrn.Text = "BRN — agence DU LOT :"
        '
        'txtBrn
        '
        Me.txtBrn.Location = New System.Drawing.Point(324, 55)
        Me.txtBrn.Name = "txtBrn"
        Me.txtBrn.Size = New System.Drawing.Size(160, 20)
        Me.txtBrn.TabIndex = 3
        '
        'lblSrccode
        '
        Me.lblSrccode.Location = New System.Drawing.Point(18, 92)
        Me.lblSrccode.Name = "lblSrccode"
        Me.lblSrccode.Size = New System.Drawing.Size(300, 20)
        Me.lblSrccode.TabIndex = 4
        Me.lblSrccode.Text = "SRCCODE — origine du lot :"
        '
        'txtSrccode
        '
        Me.txtSrccode.Location = New System.Drawing.Point(324, 89)
        Me.txtSrccode.Name = "txtSrccode"
        Me.txtSrccode.Size = New System.Drawing.Size(160, 20)
        Me.txtSrccode.TabIndex = 5
        '
        'lblTxncdDebit
        '
        Me.lblTxncdDebit.Location = New System.Drawing.Point(18, 126)
        Me.lblTxncdDebit.Name = "lblTxncdDebit"
        Me.lblTxncdDebit.Size = New System.Drawing.Size(300, 20)
        Me.lblTxncdDebit.TabIndex = 6
        Me.lblTxncdDebit.Text = "TXNCD — code transaction d'un DÉBIT :"
        '
        'txtTxncdDebit
        '
        Me.txtTxncdDebit.Location = New System.Drawing.Point(324, 123)
        Me.txtTxncdDebit.Name = "txtTxncdDebit"
        Me.txtTxncdDebit.Size = New System.Drawing.Size(160, 20)
        Me.txtTxncdDebit.TabIndex = 7
        '
        'lblTxncdCredit
        '
        Me.lblTxncdCredit.Location = New System.Drawing.Point(18, 160)
        Me.lblTxncdCredit.Name = "lblTxncdCredit"
        Me.lblTxncdCredit.Size = New System.Drawing.Size(300, 20)
        Me.lblTxncdCredit.TabIndex = 8
        Me.lblTxncdCredit.Text = "TXNCD — code transaction d'un CRÉDIT :"
        '
        'txtTxncdCredit
        '
        Me.txtTxncdCredit.Location = New System.Drawing.Point(324, 157)
        Me.txtTxncdCredit.Name = "txtTxncdCredit"
        Me.txtTxncdCredit.Size = New System.Drawing.Size(160, 20)
        Me.txtTxncdCredit.TabIndex = 9
        '
        'lblAgenceSiege
        '
        Me.lblAgenceSiege.Location = New System.Drawing.Point(18, 194)
        Me.lblAgenceSiege.Name = "lblAgenceSiege"
        Me.lblAgenceSiege.Size = New System.Drawing.Size(300, 20)
        Me.lblAgenceSiege.TabIndex = 10
        Me.lblAgenceSiege.Text = "ACBRN — agence du SIÈGE :"
        '
        'txtAgenceSiege
        '
        Me.txtAgenceSiege.Location = New System.Drawing.Point(324, 191)
        Me.txtAgenceSiege.Name = "txtAgenceSiege"
        Me.txtAgenceSiege.Size = New System.Drawing.Size(160, 20)
        Me.txtAgenceSiege.TabIndex = 11
        '
        'lblAvertissementBrn
        '
        Me.lblAvertissementBrn.ForeColor = System.Drawing.SystemColors.GrayText
        Me.lblAvertissementBrn.Location = New System.Drawing.Point(18, 230)
        Me.lblAvertissementBrn.Name = "lblAvertissementBrn"
        Me.lblAvertissementBrn.Size = New System.Drawing.Size(850, 74)
        Me.lblAvertissementBrn.TabIndex = 12
        Me.lblAvertissementBrn.Text = "BRN et ACBRN NE SONT PAS LA MÊME CHOSE. BRN désigne l'agence du LOT, portée à l'i" &
            "dentique par toutes ses lignes. ACBRN désigne l'agence de CHAQUE écriture : celle" &
            " du point de vente qui l'a produite, ou celle du siège pour les comptes de l'ongl" &
            "et voisin et pour les lignes qui n'appartiennent à aucun point de vente — l'écart" &
            " d'arrondi, les écritures de change."
        '
        'btnDefaut
        '
        Me.btnDefaut.Location = New System.Drawing.Point(18, 318)
        Me.btnDefaut.Name = "btnDefaut"
        Me.btnDefaut.Size = New System.Drawing.Size(280, 32)
        Me.btnDefaut.TabIndex = 13
        Me.btnDefaut.Text = "Rétablir les codes par défaut"
        Me.btnDefaut.UseVisualStyleBackColor = True
        '
        'ongComptes
        '
        Me.ongComptes.Controls.Add(Me.lblAideComptes)
        Me.ongComptes.Controls.Add(Me.dgvComptes)
        Me.ongComptes.Location = New System.Drawing.Point(4, 22)
        Me.ongComptes.Name = "ongComptes"
        Me.ongComptes.Padding = New System.Windows.Forms.Padding(3)
        Me.ongComptes.Size = New System.Drawing.Size(888, 414)
        Me.ongComptes.TabIndex = 1
        Me.ongComptes.Text = "Comptes rattachés au siège"
        Me.ongComptes.UseVisualStyleBackColor = True
        '
        'lblAideComptes
        '
        Me.lblAideComptes.ForeColor = System.Drawing.SystemColors.GrayText
        Me.lblAideComptes.Location = New System.Drawing.Point(18, 14)
        Me.lblAideComptes.Name = "lblAideComptes"
        Me.lblAideComptes.Size = New System.Drawing.Size(850, 56)
        Me.lblAideComptes.TabIndex = 0
        Me.lblAideComptes.Text = "Ces comptes portent l'agence du siège en ACBRN, quel que soit le point de vente q" &
            "ui les a mouvementés. CE N'EST PAS UN PLAN COMPTABLE : aucune écriture ne leur es" &
            "t passée du fait de cette liste, elle ne décide que de leur agence. Une liste vid" &
            "e est acceptée."
        '
        'dgvComptes
        '
        Me.dgvComptes.AllowUserToAddRows = True
        Me.dgvComptes.AllowUserToDeleteRows = True
        Me.dgvComptes.AllowUserToResizeRows = False
        Me.dgvComptes.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvComptes.Location = New System.Drawing.Point(18, 78)
        Me.dgvComptes.MultiSelect = False
        Me.dgvComptes.Name = "dgvComptes"
        Me.dgvComptes.RowHeadersWidth = 28
        Me.dgvComptes.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect
        Me.dgvComptes.Size = New System.Drawing.Size(500, 310)
        Me.dgvComptes.TabIndex = 1
        '
        'ongJournal
        '
        Me.ongJournal.Controls.Add(Me.lblJournal)
        Me.ongJournal.Controls.Add(Me.dgvJournal)
        Me.ongJournal.Location = New System.Drawing.Point(4, 22)
        Me.ongJournal.Name = "ongJournal"
        Me.ongJournal.Padding = New System.Windows.Forms.Padding(3)
        Me.ongJournal.Size = New System.Drawing.Size(888, 414)
        Me.ongJournal.TabIndex = 2
        Me.ongJournal.Text = "Historique des modifications"
        Me.ongJournal.UseVisualStyleBackColor = True
        '
        'lblJournal
        '
        Me.lblJournal.ForeColor = System.Drawing.SystemColors.GrayText
        Me.lblJournal.Location = New System.Drawing.Point(18, 14)
        Me.lblJournal.Name = "lblJournal"
        Me.lblJournal.Size = New System.Drawing.Size(850, 34)
        Me.lblJournal.TabIndex = 0
        Me.lblJournal.Text = "Qui a changé quoi, quand, depuis quel poste, et ce qu'il y avait avant."
        '
        'dgvJournal
        '
        Me.dgvJournal.AllowUserToAddRows = False
        Me.dgvJournal.AllowUserToDeleteRows = False
        Me.dgvJournal.AllowUserToResizeRows = False
        Me.dgvJournal.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvJournal.Location = New System.Drawing.Point(18, 54)
        Me.dgvJournal.MultiSelect = False
        Me.dgvJournal.Name = "dgvJournal"
        Me.dgvJournal.ReadOnly = True
        Me.dgvJournal.RowHeadersWidth = 28
        Me.dgvJournal.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvJournal.Size = New System.Drawing.Size(850, 334)
        Me.dgvJournal.TabIndex = 1
        '
        'lblApercuTitre
        '
        Me.lblApercuTitre.Location = New System.Drawing.Point(12, 540)
        Me.lblApercuTitre.Name = "lblApercuTitre"
        Me.lblApercuTitre.Size = New System.Drawing.Size(896, 18)
        Me.lblApercuTitre.TabIndex = 3
        Me.lblApercuTitre.Text = "Aperçu — une ligne du fichier telle qu'elle sortira :"
        '
        'lblApercu
        '
        Me.lblApercu.Font = New System.Drawing.Font("Consolas", 9.0!)
        Me.lblApercu.Location = New System.Drawing.Point(24, 560)
        Me.lblApercu.Name = "lblApercu"
        Me.lblApercu.Size = New System.Drawing.Size(884, 36)
        Me.lblApercu.TabIndex = 4
        '
        'lblDerniereModification
        '
        Me.lblDerniereModification.ForeColor = System.Drawing.SystemColors.GrayText
        Me.lblDerniereModification.Location = New System.Drawing.Point(12, 600)
        Me.lblDerniereModification.Name = "lblDerniereModification"
        Me.lblDerniereModification.Size = New System.Drawing.Size(676, 18)
        Me.lblDerniereModification.TabIndex = 5
        '
        'lblStatut
        '
        Me.lblStatut.ForeColor = System.Drawing.SystemColors.GrayText
        Me.lblStatut.Location = New System.Drawing.Point(12, 622)
        Me.lblStatut.Name = "lblStatut"
        Me.lblStatut.Size = New System.Drawing.Size(676, 18)
        Me.lblStatut.TabIndex = 6
        '
        'btnEnregistrer
        '
        Me.btnEnregistrer.Location = New System.Drawing.Point(698, 612)
        Me.btnEnregistrer.Name = "btnEnregistrer"
        Me.btnEnregistrer.Size = New System.Drawing.Size(110, 32)
        Me.btnEnregistrer.TabIndex = 7
        Me.btnEnregistrer.Text = "Enregistrer"
        Me.btnEnregistrer.UseVisualStyleBackColor = True
        '
        'btnFermer
        '
        Me.btnFermer.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnFermer.Location = New System.Drawing.Point(816, 612)
        Me.btnFermer.Name = "btnFermer"
        Me.btnFermer.Size = New System.Drawing.Size(92, 32)
        Me.btnFermer.TabIndex = 8
        Me.btnFermer.Text = "Fermer"
        Me.btnFermer.UseVisualStyleBackColor = True
        '
        'FrmFormatCoreBanking
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnFermer
        Me.ClientSize = New System.Drawing.Size(920, 658)
        Me.Controls.Add(Me.btnFermer)
        Me.Controls.Add(Me.btnEnregistrer)
        Me.Controls.Add(Me.lblStatut)
        Me.Controls.Add(Me.lblDerniereModification)
        Me.Controls.Add(Me.lblApercu)
        Me.Controls.Add(Me.lblApercuTitre)
        Me.Controls.Add(Me.ongletsFormat)
        Me.Controls.Add(Me.lblSousTitre)
        Me.Controls.Add(Me.lblTitre)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FrmFormatCoreBanking"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Format du fichier core banking"
        Me.ongCodes.ResumeLayout(False)
        Me.ongCodes.PerformLayout()
        Me.ongComptes.ResumeLayout(False)
        Me.ongJournal.ResumeLayout(False)
        Me.ongletsFormat.ResumeLayout(False)
        CType(Me.dgvComptes, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvJournal, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents lblTitre As System.Windows.Forms.Label
    Friend WithEvents lblSousTitre As System.Windows.Forms.Label
    Friend WithEvents ongletsFormat As System.Windows.Forms.TabControl
    Friend WithEvents ongCodes As System.Windows.Forms.TabPage
    Friend WithEvents lblDetbsjrnl As System.Windows.Forms.Label
    Friend WithEvents txtDetbsjrnl As System.Windows.Forms.TextBox
    Friend WithEvents lblBrn As System.Windows.Forms.Label
    Friend WithEvents txtBrn As System.Windows.Forms.TextBox
    Friend WithEvents lblSrccode As System.Windows.Forms.Label
    Friend WithEvents txtSrccode As System.Windows.Forms.TextBox
    Friend WithEvents lblTxncdDebit As System.Windows.Forms.Label
    Friend WithEvents txtTxncdDebit As System.Windows.Forms.TextBox
    Friend WithEvents lblTxncdCredit As System.Windows.Forms.Label
    Friend WithEvents txtTxncdCredit As System.Windows.Forms.TextBox
    Friend WithEvents lblAgenceSiege As System.Windows.Forms.Label
    Friend WithEvents txtAgenceSiege As System.Windows.Forms.TextBox
    Friend WithEvents lblAvertissementBrn As System.Windows.Forms.Label
    Friend WithEvents btnDefaut As System.Windows.Forms.Button
    Friend WithEvents ongComptes As System.Windows.Forms.TabPage
    Friend WithEvents lblAideComptes As System.Windows.Forms.Label
    Friend WithEvents dgvComptes As System.Windows.Forms.DataGridView
    Friend WithEvents ongJournal As System.Windows.Forms.TabPage
    Friend WithEvents lblJournal As System.Windows.Forms.Label
    Friend WithEvents dgvJournal As System.Windows.Forms.DataGridView
    Friend WithEvents lblApercuTitre As System.Windows.Forms.Label
    Friend WithEvents lblApercu As System.Windows.Forms.Label
    Friend WithEvents lblDerniereModification As System.Windows.Forms.Label
    Friend WithEvents lblStatut As System.Windows.Forms.Label
    Friend WithEvents btnEnregistrer As System.Windows.Forms.Button
    Friend WithEvents btnFermer As System.Windows.Forms.Button

End Class
