Option Strict On
Option Explicit On

Partial Class FrmDiagnostic
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
        Me.txtDiagnostic = New System.Windows.Forms.TextBox()
        Me.lblAide = New System.Windows.Forms.Label()
        Me.btnCopier = New System.Windows.Forms.Button()
        Me.btnFermer = New System.Windows.Forms.Button()
        Me.panelTitre.SuspendLayout()
        Me.SuspendLayout()
        '
        'panelTitre
        '
        Me.panelTitre.BackColor = System.Drawing.Color.Black
        Me.panelTitre.Controls.Add(Me.lblTitre)
        Me.panelTitre.Location = New System.Drawing.Point(12, 12)
        Me.panelTitre.Size = New System.Drawing.Size(646, 56)
        Me.panelTitre.Name = "panelTitre"
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
        Me.lblTitre.Size = New System.Drawing.Size(646, 56)
        Me.lblTitre.TabIndex = 0
        Me.lblTitre.Text = "Diagnostic"
        Me.lblTitre.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtDiagnostic
        '
        ' Une police à chasse fixe : le T-SQL du diagnostic s'aligne, et se relit.
        Me.txtDiagnostic.Font = New System.Drawing.Font("Consolas", 9.75!)
        Me.txtDiagnostic.Location = New System.Drawing.Point(12, 80)
        Me.txtDiagnostic.Multiline = True
        Me.txtDiagnostic.Name = "txtDiagnostic"
        Me.txtDiagnostic.ReadOnly = True
        Me.txtDiagnostic.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtDiagnostic.Size = New System.Drawing.Size(646, 330)
        Me.txtDiagnostic.TabIndex = 1
        '
        'lblAide
        '
        Me.lblAide.ForeColor = System.Drawing.SystemColors.GrayText
        Me.lblAide.Location = New System.Drawing.Point(12, 420)
        Me.lblAide.Name = "lblAide"
        Me.lblAide.Size = New System.Drawing.Size(450, 44)
        Me.lblAide.TabIndex = 2
        Me.lblAide.Text = "« Copier » met ce texte dans le presse-papiers : il se colle tel quel dans un courriel à l'informatique."
        '
        'btnCopier
        '
        Me.btnCopier.Location = New System.Drawing.Point(468, 424)
        Me.btnCopier.Name = "btnCopier"
        Me.btnCopier.Size = New System.Drawing.Size(90, 32)
        Me.btnCopier.TabIndex = 3
        Me.btnCopier.Text = "Copier"
        Me.btnCopier.UseVisualStyleBackColor = True
        '
        'btnFermer
        '
        Me.btnFermer.Location = New System.Drawing.Point(568, 424)
        Me.btnFermer.Name = "btnFermer"
        Me.btnFermer.Size = New System.Drawing.Size(90, 32)
        Me.btnFermer.TabIndex = 4
        Me.btnFermer.Text = "Fermer"
        Me.btnFermer.UseVisualStyleBackColor = True
        '
        'FrmDiagnostic
        '
        Me.AcceptButton = Me.btnFermer
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnFermer
        Me.ClientSize = New System.Drawing.Size(670, 470)
        Me.Controls.Add(Me.btnFermer)
        Me.Controls.Add(Me.btnCopier)
        Me.Controls.Add(Me.lblAide)
        Me.Controls.Add(Me.txtDiagnostic)
        Me.Controls.Add(Me.panelTitre)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FrmDiagnostic"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Diagnostic"
        Me.panelTitre.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()
    End Sub

    Friend WithEvents panelTitre As System.Windows.Forms.Panel
    Friend WithEvents lblTitre As System.Windows.Forms.Label
    Friend WithEvents txtDiagnostic As System.Windows.Forms.TextBox
    Friend WithEvents lblAide As System.Windows.Forms.Label
    Friend WithEvents btnCopier As System.Windows.Forms.Button
    Friend WithEvents btnFermer As System.Windows.Forms.Button

End Class
