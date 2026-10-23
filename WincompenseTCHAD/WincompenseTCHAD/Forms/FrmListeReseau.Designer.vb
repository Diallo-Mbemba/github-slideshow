Option Strict On
Option Explicit On

Partial Class FrmListeReseau
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
        Me.grpPortee = New System.Windows.Forms.GroupBox()
        Me.rdoTout = New System.Windows.Forms.RadioButton()
        Me.rdoGroupe = New System.Windows.Forms.RadioButton()
        Me.cboGroupe = New System.Windows.Forms.ComboBox()
        Me.grpPeriode = New System.Windows.Forms.GroupBox()
        Me.chkPeriode = New System.Windows.Forms.CheckBox()
        Me.dtpDebut = New System.Windows.Forms.DateTimePicker()
        Me.lblEt = New System.Windows.Forms.Label()
        Me.dtpFin = New System.Windows.Forms.DateTimePicker()
        Me.btnAfficher = New System.Windows.Forms.Button()
        Me.dgvReseau = New System.Windows.Forms.DataGridView()
        Me.lblTotal = New System.Windows.Forms.Label()
        Me.btnImprimer = New System.Windows.Forms.Button()
        Me.btnExporter = New System.Windows.Forms.Button()
        Me.btnFermer = New System.Windows.Forms.Button()
        Me.sfdEtat = New System.Windows.Forms.SaveFileDialog()
        Me.grpPortee.SuspendLayout()
        Me.grpPeriode.SuspendLayout()
        CType(Me.dgvReseau, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'grpPortee
        '
        Me.grpPortee.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grpPortee.Controls.Add(Me.rdoTout)
        Me.grpPortee.Controls.Add(Me.rdoGroupe)
        Me.grpPortee.Controls.Add(Me.cboGroupe)
        Me.grpPortee.Location = New System.Drawing.Point(12, 12)
        Me.grpPortee.Name = "grpPortee"
        Me.grpPortee.Size = New System.Drawing.Size(560, 104)
        Me.grpPortee.TabIndex = 0
        Me.grpPortee.TabStop = False
        Me.grpPortee.Text = "Que faut-il imprimer ?"
        '
        'rdoTout
        '
        Me.rdoTout.Checked = True
        Me.rdoTout.Location = New System.Drawing.Point(16, 26)
        Me.rdoTout.Name = "rdoTout"
        Me.rdoTout.Size = New System.Drawing.Size(530, 22)
        Me.rdoTout.TabIndex = 0
        Me.rdoTout.TabStop = True
        Me.rdoTout.Text = "Tout le réseau : groupes statistiques, sous-agents et agences propres"
        Me.rdoTout.UseVisualStyleBackColor = True
        '
        'rdoGroupe
        '
        Me.rdoGroupe.Location = New System.Drawing.Point(16, 58)
        Me.rdoGroupe.Name = "rdoGroupe"
        Me.rdoGroupe.Size = New System.Drawing.Size(160, 22)
        Me.rdoGroupe.TabIndex = 1
        Me.rdoGroupe.Text = "Un groupe statistique :"
        Me.rdoGroupe.UseVisualStyleBackColor = True
        '
        'cboGroupe
        '
        Me.cboGroupe.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboGroupe.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboGroupe.FormattingEnabled = True
        Me.cboGroupe.Location = New System.Drawing.Point(182, 56)
        Me.cboGroupe.Name = "cboGroupe"
        Me.cboGroupe.Size = New System.Drawing.Size(360, 24)
        Me.cboGroupe.TabIndex = 2
        '
        'grpPeriode
        '
        Me.grpPeriode.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grpPeriode.Controls.Add(Me.chkPeriode)
        Me.grpPeriode.Controls.Add(Me.dtpDebut)
        Me.grpPeriode.Controls.Add(Me.lblEt)
        Me.grpPeriode.Controls.Add(Me.dtpFin)
        Me.grpPeriode.Location = New System.Drawing.Point(584, 12)
        Me.grpPeriode.Name = "grpPeriode"
        Me.grpPeriode.Size = New System.Drawing.Size(344, 104)
        Me.grpPeriode.TabIndex = 1
        Me.grpPeriode.TabStop = False
        Me.grpPeriode.Text = "Date de création"
        '
        'chkPeriode
        '
        Me.chkPeriode.Location = New System.Drawing.Point(14, 26)
        Me.chkPeriode.Name = "chkPeriode"
        Me.chkPeriode.Size = New System.Drawing.Size(320, 22)
        Me.chkPeriode.TabIndex = 0
        Me.chkPeriode.Text = "Seulement les points de vente créés entre le"
        Me.chkPeriode.UseVisualStyleBackColor = True
        '
        'dtpDebut
        '
        Me.dtpDebut.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtpDebut.Location = New System.Drawing.Point(14, 58)
        Me.dtpDebut.Name = "dtpDebut"
        Me.dtpDebut.Size = New System.Drawing.Size(120, 22)
        Me.dtpDebut.TabIndex = 1
        '
        'lblEt
        '
        Me.lblEt.Location = New System.Drawing.Point(142, 61)
        Me.lblEt.Name = "lblEt"
        Me.lblEt.Size = New System.Drawing.Size(34, 20)
        Me.lblEt.TabIndex = 2
        Me.lblEt.Text = "et le"
        '
        'dtpFin
        '
        Me.dtpFin.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtpFin.Location = New System.Drawing.Point(180, 58)
        Me.dtpFin.Name = "dtpFin"
        Me.dtpFin.Size = New System.Drawing.Size(120, 22)
        Me.dtpFin.TabIndex = 3
        '
        'btnAfficher
        '
        Me.btnAfficher.Location = New System.Drawing.Point(12, 128)
        Me.btnAfficher.Name = "btnAfficher"
        Me.btnAfficher.Size = New System.Drawing.Size(140, 30)
        Me.btnAfficher.TabIndex = 2
        Me.btnAfficher.Text = "Afficher"
        Me.btnAfficher.UseVisualStyleBackColor = True
        '
        'dgvReseau
        '
        Me.dgvReseau.AllowUserToAddRows = False
        Me.dgvReseau.AllowUserToDeleteRows = False
        Me.dgvReseau.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvReseau.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvReseau.Location = New System.Drawing.Point(12, 166)
        Me.dgvReseau.MultiSelect = False
        Me.dgvReseau.Name = "dgvReseau"
        Me.dgvReseau.ReadOnly = True
        Me.dgvReseau.RowHeadersWidth = 25
        Me.dgvReseau.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvReseau.Size = New System.Drawing.Size(916, 340)
        Me.dgvReseau.TabIndex = 3
        '
        'lblTotal
        '
        Me.lblTotal.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblTotal.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold)
        Me.lblTotal.Location = New System.Drawing.Point(12, 514)
        Me.lblTotal.Name = "lblTotal"
        Me.lblTotal.Size = New System.Drawing.Size(916, 36)
        Me.lblTotal.TabIndex = 4
        Me.lblTotal.Text = ""
        '
        'btnImprimer
        '
        Me.btnImprimer.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnImprimer.Location = New System.Drawing.Point(12, 556)
        Me.btnImprimer.Name = "btnImprimer"
        Me.btnImprimer.Size = New System.Drawing.Size(176, 32)
        Me.btnImprimer.TabIndex = 5
        Me.btnImprimer.Text = "Imprimer (PDF)"
        Me.btnImprimer.UseVisualStyleBackColor = True
        '
        'btnExporter
        '
        Me.btnExporter.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnExporter.Location = New System.Drawing.Point(196, 556)
        Me.btnExporter.Name = "btnExporter"
        Me.btnExporter.Size = New System.Drawing.Size(200, 32)
        Me.btnExporter.TabIndex = 6
        Me.btnExporter.Text = "Exporter vers Excel"
        Me.btnExporter.UseVisualStyleBackColor = True
        '
        'btnFermer
        '
        Me.btnFermer.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFermer.Location = New System.Drawing.Point(818, 556)
        Me.btnFermer.Name = "btnFermer"
        Me.btnFermer.Size = New System.Drawing.Size(110, 32)
        Me.btnFermer.TabIndex = 7
        Me.btnFermer.Text = "Fermer"
        Me.btnFermer.UseVisualStyleBackColor = True
        '
        'FrmListeReseau
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnFermer
        Me.ClientSize = New System.Drawing.Size(940, 600)
        Me.Controls.Add(Me.grpPortee)
        Me.Controls.Add(Me.grpPeriode)
        Me.Controls.Add(Me.btnAfficher)
        Me.Controls.Add(Me.dgvReseau)
        Me.Controls.Add(Me.lblTotal)
        Me.Controls.Add(Me.btnImprimer)
        Me.Controls.Add(Me.btnExporter)
        Me.Controls.Add(Me.btnFermer)
        Me.MinimumSize = New System.Drawing.Size(956, 640)
        Me.Name = "FrmListeReseau"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Liste du réseau"
        Me.grpPortee.ResumeLayout(False)
        Me.grpPeriode.ResumeLayout(False)
        CType(Me.dgvReseau, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
    End Sub

    Friend WithEvents grpPortee As System.Windows.Forms.GroupBox
    Friend WithEvents rdoTout As System.Windows.Forms.RadioButton
    Friend WithEvents rdoGroupe As System.Windows.Forms.RadioButton
    Friend WithEvents cboGroupe As System.Windows.Forms.ComboBox
    Friend WithEvents grpPeriode As System.Windows.Forms.GroupBox
    Friend WithEvents chkPeriode As System.Windows.Forms.CheckBox
    Friend WithEvents dtpDebut As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblEt As System.Windows.Forms.Label
    Friend WithEvents dtpFin As System.Windows.Forms.DateTimePicker
    Friend WithEvents btnAfficher As System.Windows.Forms.Button
    Friend WithEvents dgvReseau As System.Windows.Forms.DataGridView
    Friend WithEvents lblTotal As System.Windows.Forms.Label
    Friend WithEvents btnImprimer As System.Windows.Forms.Button
    Friend WithEvents btnExporter As System.Windows.Forms.Button
    Friend WithEvents btnFermer As System.Windows.Forms.Button
    Friend WithEvents sfdEtat As System.Windows.Forms.SaveFileDialog

End Class
