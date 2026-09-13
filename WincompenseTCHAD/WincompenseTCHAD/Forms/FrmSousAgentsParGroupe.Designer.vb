Option Strict On
Option Explicit On

Partial Class FrmSousAgentsParGroupe
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
        Me.lblFiltre = New System.Windows.Forms.Label()
        Me.lblNombre = New System.Windows.Forms.Label()
        Me.lblStatut = New System.Windows.Forms.Label()
        Me.lblGroupes = New System.Windows.Forms.Label()
        Me.lblDetail = New System.Windows.Forms.Label()
        Me.cboGroupe = New System.Windows.Forms.ComboBox()
        Me.btnActualiser = New System.Windows.Forms.Button()
        Me.btnExporter = New System.Windows.Forms.Button()
        Me.btnFermer = New System.Windows.Forms.Button()
        Me.dgvGroupes = New System.Windows.Forms.DataGridView()
        Me.dgvSousAgents = New System.Windows.Forms.DataGridView()
        CType(Me.dgvGroupes, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvSousAgents, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.sfdExport = New System.Windows.Forms.SaveFileDialog()
        Me.SuspendLayout()
        '
        'lblFiltre
        '
        Me.lblFiltre.Location = New System.Drawing.Point(12, 16)
        Me.lblFiltre.Name = "lblFiltre"
        Me.lblFiltre.Size = New System.Drawing.Size(110, 20)
        Me.lblFiltre.TabIndex = 0
        Me.lblFiltre.Text = "Groupe affiché :"
        '
        'cboGroupe
        '
        Me.cboGroupe.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboGroupe.FormattingEnabled = True
        Me.cboGroupe.Location = New System.Drawing.Point(125, 12)
        Me.cboGroupe.Name = "cboGroupe"
        Me.cboGroupe.Size = New System.Drawing.Size(280, 24)
        Me.cboGroupe.TabIndex = 1
        '
        'btnActualiser
        '
        Me.btnActualiser.Location = New System.Drawing.Point(415, 11)
        Me.btnActualiser.Name = "btnActualiser"
        Me.btnActualiser.Size = New System.Drawing.Size(110, 26)
        Me.btnActualiser.TabIndex = 2
        Me.btnActualiser.Text = "Actualiser"
        Me.btnActualiser.UseVisualStyleBackColor = True
        '
        'lblNombre
        '
        Me.lblNombre.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblNombre.Location = New System.Drawing.Point(568, 16)
        Me.lblNombre.Name = "lblNombre"
        Me.lblNombre.Size = New System.Drawing.Size(360, 20)
        Me.lblNombre.TabIndex = 3
        Me.lblNombre.Text = ""
        Me.lblNombre.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblGroupes
        '
        Me.lblGroupes.Location = New System.Drawing.Point(12, 46)
        Me.lblGroupes.Name = "lblGroupes"
        Me.lblGroupes.Size = New System.Drawing.Size(916, 18)
        Me.lblGroupes.TabIndex = 4
        Me.lblGroupes.Text = "Groupes statistiques (cliquez un groupe pour n'afficher que ses sous-agents)"
        '
        'dgvGroupes
        '
        Me.dgvGroupes.AllowUserToAddRows = False
        Me.dgvGroupes.AllowUserToDeleteRows = False
        Me.dgvGroupes.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvGroupes.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvGroupes.Location = New System.Drawing.Point(12, 66)
        Me.dgvGroupes.MultiSelect = False
        Me.dgvGroupes.Name = "dgvGroupes"
        Me.dgvGroupes.ReadOnly = True
        Me.dgvGroupes.RowHeadersWidth = 25
        Me.dgvGroupes.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvGroupes.Size = New System.Drawing.Size(916, 170)
        Me.dgvGroupes.TabIndex = 5
        '
        'lblDetail
        '
        Me.lblDetail.Location = New System.Drawing.Point(12, 246)
        Me.lblDetail.Name = "lblDetail"
        Me.lblDetail.Size = New System.Drawing.Size(916, 18)
        Me.lblDetail.TabIndex = 6
        Me.lblDetail.Text = "Sous-agents"
        '
        'dgvSousAgents
        '
        Me.dgvSousAgents.AllowUserToAddRows = False
        Me.dgvSousAgents.AllowUserToDeleteRows = False
        Me.dgvSousAgents.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvSousAgents.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvSousAgents.Location = New System.Drawing.Point(12, 266)
        Me.dgvSousAgents.MultiSelect = False
        Me.dgvSousAgents.Name = "dgvSousAgents"
        Me.dgvSousAgents.ReadOnly = True
        Me.dgvSousAgents.RowHeadersWidth = 25
        Me.dgvSousAgents.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvSousAgents.Size = New System.Drawing.Size(916, 300)
        Me.dgvSousAgents.TabIndex = 7
        '
        'btnExporter
        '
        Me.btnExporter.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnExporter.Location = New System.Drawing.Point(12, 578)
        Me.btnExporter.Name = "btnExporter"
        Me.btnExporter.Size = New System.Drawing.Size(220, 32)
        Me.btnExporter.TabIndex = 8
        Me.btnExporter.Text = "Exporter vers Excel..."
        Me.btnExporter.UseVisualStyleBackColor = True
        '
        'sfdExport
        '
        Me.sfdExport.DefaultExt = "xlsx"
        Me.sfdExport.Filter = "Classeur Excel (*.xlsx)|*.xlsx"
        Me.sfdExport.Title = "Exporter la liste des sous-agents par groupe"
        '
        'btnFermer
        '
        Me.btnFermer.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFermer.Location = New System.Drawing.Point(808, 578)
        Me.btnFermer.Name = "btnFermer"
        Me.btnFermer.Size = New System.Drawing.Size(120, 32)
        Me.btnFermer.TabIndex = 9
        Me.btnFermer.Text = "Fermer"
        Me.btnFermer.UseVisualStyleBackColor = True
        '
        'lblStatut
        '
        Me.lblStatut.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblStatut.ForeColor = System.Drawing.SystemColors.GrayText
        Me.lblStatut.Location = New System.Drawing.Point(12, 586)
        Me.lblStatut.Name = "lblStatut"
        Me.lblStatut.Size = New System.Drawing.Size(784, 20)
        Me.lblStatut.TabIndex = 10
        Me.lblStatut.Text = ""
        '
        'FrmSousAgentsParGroupe
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnFermer
        Me.ClientSize = New System.Drawing.Size(940, 620)
        Me.Controls.Add(Me.lblFiltre)
        Me.Controls.Add(Me.cboGroupe)
        Me.Controls.Add(Me.btnActualiser)
        Me.Controls.Add(Me.lblNombre)
        Me.Controls.Add(Me.lblGroupes)
        Me.Controls.Add(Me.dgvGroupes)
        Me.Controls.Add(Me.lblDetail)
        Me.Controls.Add(Me.dgvSousAgents)
        Me.Controls.Add(Me.btnExporter)
        Me.Controls.Add(Me.btnFermer)
        Me.Controls.Add(Me.lblStatut)
        Me.MinimumSize = New System.Drawing.Size(956, 660)
        Me.Name = "FrmSousAgentsParGroupe"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Sous-agents par groupe statistique"
        CType(Me.dgvGroupes, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvSousAgents, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
    End Sub

    Friend WithEvents lblFiltre As System.Windows.Forms.Label
    Friend WithEvents lblNombre As System.Windows.Forms.Label
    Friend WithEvents lblStatut As System.Windows.Forms.Label
    Friend WithEvents lblGroupes As System.Windows.Forms.Label
    Friend WithEvents lblDetail As System.Windows.Forms.Label
    Friend WithEvents cboGroupe As System.Windows.Forms.ComboBox
    Friend WithEvents btnActualiser As System.Windows.Forms.Button
    Friend WithEvents btnExporter As System.Windows.Forms.Button
    Friend WithEvents sfdExport As System.Windows.Forms.SaveFileDialog
    Friend WithEvents btnFermer As System.Windows.Forms.Button
    Friend WithEvents dgvGroupes As System.Windows.Forms.DataGridView
    Friend WithEvents dgvSousAgents As System.Windows.Forms.DataGridView

End Class
