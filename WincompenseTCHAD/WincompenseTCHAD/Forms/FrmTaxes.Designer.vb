Option Strict On
Option Explicit On

Partial Class FrmTaxes
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
        Me.lblIntro = New System.Windows.Forms.Label()
        Me.panelVerrou = New System.Windows.Forms.Panel()
        Me.lblVerrou = New System.Windows.Forms.Label()
        Me.dgvTaxes = New System.Windows.Forms.DataGridView()
        Me.lblFormule = New System.Windows.Forms.Label()
        Me.lblNotes = New System.Windows.Forms.Label()
        Me.btnFermer = New System.Windows.Forms.Button()
        Me.panelTitre.SuspendLayout()
        Me.panelVerrou.SuspendLayout()
        CType(Me.dgvTaxes, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'panelTitre
        '
        Me.panelTitre.BackColor = System.Drawing.Color.Black
        Me.panelTitre.Controls.Add(Me.lblTitre)
        Me.panelTitre.Location = New System.Drawing.Point(12, 12)
        Me.panelTitre.Name = "panelTitre"
        Me.panelTitre.Size = New System.Drawing.Size(776, 62)
        Me.panelTitre.TabIndex = 0
        '
        'lblTitre
        '
        Me.lblTitre.BackColor = System.Drawing.Color.Black
        Me.lblTitre.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblTitre.Font = New System.Drawing.Font("Segoe UI", 24.0!, System.Drawing.FontStyle.Bold)
        Me.lblTitre.ForeColor = System.Drawing.Color.Yellow
        Me.lblTitre.Location = New System.Drawing.Point(0, 0)
        Me.lblTitre.Name = "lblTitre"
        Me.lblTitre.Size = New System.Drawing.Size(776, 62)
        Me.lblTitre.TabIndex = 0
        Me.lblTitre.Text = "Taxes et barème"
        Me.lblTitre.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblIntro
        '
        Me.lblIntro.ForeColor = System.Drawing.SystemColors.ControlDarkDark
        Me.lblIntro.Location = New System.Drawing.Point(12, 82)
        Me.lblIntro.Name = "lblIntro"
        Me.lblIntro.Size = New System.Drawing.Size(776, 34)
        Me.lblIntro.TabIndex = 1
        '
        'panelVerrou
        '
        Me.panelVerrou.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.panelVerrou.Controls.Add(Me.lblVerrou)
        Me.panelVerrou.Location = New System.Drawing.Point(12, 120)
        Me.panelVerrou.Name = "panelVerrou"
        Me.panelVerrou.Padding = New System.Windows.Forms.Padding(10, 6, 10, 6)
        Me.panelVerrou.Size = New System.Drawing.Size(776, 58)
        Me.panelVerrou.TabIndex = 2
        '
        'lblVerrou
        '
        Me.lblVerrou.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblVerrou.Location = New System.Drawing.Point(10, 6)
        Me.lblVerrou.Name = "lblVerrou"
        Me.lblVerrou.Size = New System.Drawing.Size(754, 44)
        Me.lblVerrou.TabIndex = 0
        '
        'dgvTaxes
        '
        Me.dgvTaxes.AllowUserToAddRows = False
        Me.dgvTaxes.AllowUserToDeleteRows = False
        Me.dgvTaxes.AllowUserToResizeRows = False
        Me.dgvTaxes.BackgroundColor = System.Drawing.SystemColors.Window
        Me.dgvTaxes.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.dgvTaxes.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvTaxes.Location = New System.Drawing.Point(12, 188)
        Me.dgvTaxes.MultiSelect = False
        Me.dgvTaxes.Name = "dgvTaxes"
        Me.dgvTaxes.ReadOnly = True
        Me.dgvTaxes.RowHeadersVisible = False
        Me.dgvTaxes.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvTaxes.Size = New System.Drawing.Size(776, 196)
        Me.dgvTaxes.TabIndex = 3
        '
        'lblFormule
        '
        Me.lblFormule.Font = New System.Drawing.Font("Consolas", 9.75!)
        Me.lblFormule.Location = New System.Drawing.Point(12, 394)
        Me.lblFormule.Name = "lblFormule"
        Me.lblFormule.Size = New System.Drawing.Size(776, 22)
        Me.lblFormule.TabIndex = 4
        '
        'lblNotes
        '
        Me.lblNotes.ForeColor = System.Drawing.SystemColors.ControlDarkDark
        Me.lblNotes.Location = New System.Drawing.Point(12, 422)
        Me.lblNotes.Name = "lblNotes"
        Me.lblNotes.Size = New System.Drawing.Size(776, 86)
        Me.lblNotes.TabIndex = 5
        '
        'btnFermer
        '
        Me.btnFermer.Location = New System.Drawing.Point(688, 518)
        Me.btnFermer.Name = "btnFermer"
        Me.btnFermer.Size = New System.Drawing.Size(100, 30)
        Me.btnFermer.TabIndex = 6
        Me.btnFermer.Text = "Fermer"
        Me.btnFermer.UseVisualStyleBackColor = True
        '
        'FrmTaxes
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnFermer
        Me.ClientSize = New System.Drawing.Size(800, 560)
        Me.Controls.Add(Me.panelTitre)
        Me.Controls.Add(Me.lblIntro)
        Me.Controls.Add(Me.panelVerrou)
        Me.Controls.Add(Me.dgvTaxes)
        Me.Controls.Add(Me.lblFormule)
        Me.Controls.Add(Me.lblNotes)
        Me.Controls.Add(Me.btnFermer)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FrmTaxes"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Taxes et barème"
        Me.panelTitre.ResumeLayout(False)
        Me.panelVerrou.ResumeLayout(False)
        CType(Me.dgvTaxes, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
    End Sub

    Friend WithEvents panelTitre As System.Windows.Forms.Panel
    Friend WithEvents lblTitre As System.Windows.Forms.Label
    Friend WithEvents lblIntro As System.Windows.Forms.Label
    Friend WithEvents panelVerrou As System.Windows.Forms.Panel
    Friend WithEvents lblVerrou As System.Windows.Forms.Label
    Friend WithEvents dgvTaxes As System.Windows.Forms.DataGridView
    Friend WithEvents lblFormule As System.Windows.Forms.Label
    Friend WithEvents lblNotes As System.Windows.Forms.Label
    Friend WithEvents btnFermer As System.Windows.Forms.Button

End Class
