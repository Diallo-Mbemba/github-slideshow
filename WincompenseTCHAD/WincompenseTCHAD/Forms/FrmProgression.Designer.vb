Option Strict On
Option Explicit On

Partial Class FrmProgression
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
        Me.lblEtape = New System.Windows.Forms.Label()
        Me.barre = New System.Windows.Forms.ProgressBar()
        Me.lblCompteur = New System.Windows.Forms.Label()
        Me.lblPatience = New System.Windows.Forms.Label()
        Me.panelTitre.SuspendLayout()
        Me.SuspendLayout()
        '
        'panelTitre
        '
        Me.panelTitre.BackColor = System.Drawing.Color.Black
        Me.panelTitre.Controls.Add(Me.lblTitre)
        Me.panelTitre.Location = New System.Drawing.Point(0, 0)
        Me.panelTitre.Name = "panelTitre"
        Me.panelTitre.Size = New System.Drawing.Size(480, 44)
        Me.panelTitre.TabIndex = 0
        '
        'lblTitre
        '
        Me.lblTitre.BackColor = System.Drawing.Color.Black
        Me.lblTitre.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblTitre.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Bold)
        Me.lblTitre.ForeColor = System.Drawing.Color.Yellow
        Me.lblTitre.Location = New System.Drawing.Point(0, 0)
        Me.lblTitre.Name = "lblTitre"
        Me.lblTitre.Size = New System.Drawing.Size(480, 44)
        Me.lblTitre.TabIndex = 0
        Me.lblTitre.Text = "   Traitement en cours"
        Me.lblTitre.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblEtape
        '
        Me.lblEtape.Location = New System.Drawing.Point(16, 58)
        Me.lblEtape.Name = "lblEtape"
        Me.lblEtape.Size = New System.Drawing.Size(448, 34)
        Me.lblEtape.TabIndex = 1
        Me.lblEtape.Text = "Préparation…"
        '
        'barre
        '
        Me.barre.Location = New System.Drawing.Point(16, 96)
        Me.barre.Name = "barre"
        Me.barre.Size = New System.Drawing.Size(448, 20)
        Me.barre.TabIndex = 2
        '
        'lblCompteur
        '
        Me.lblCompteur.ForeColor = System.Drawing.SystemColors.GrayText
        Me.lblCompteur.Location = New System.Drawing.Point(16, 120)
        Me.lblCompteur.Name = "lblCompteur"
        Me.lblCompteur.Size = New System.Drawing.Size(200, 18)
        Me.lblCompteur.TabIndex = 3
        '
        'lblPatience
        '
        Me.lblPatience.ForeColor = System.Drawing.SystemColors.GrayText
        Me.lblPatience.Location = New System.Drawing.Point(216, 120)
        Me.lblPatience.Name = "lblPatience"
        Me.lblPatience.Size = New System.Drawing.Size(248, 18)
        Me.lblPatience.TabIndex = 4
        Me.lblPatience.Text = "Cette fenêtre se ferme d'elle-même."
        Me.lblPatience.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'FrmProgression
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(480, 150)
        Me.ControlBox = False
        Me.Controls.Add(Me.lblPatience)
        Me.Controls.Add(Me.lblCompteur)
        Me.Controls.Add(Me.barre)
        Me.Controls.Add(Me.lblEtape)
        Me.Controls.Add(Me.panelTitre)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FrmProgression"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Traitement en cours"
        Me.panelTitre.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents panelTitre As System.Windows.Forms.Panel
    Friend WithEvents lblTitre As System.Windows.Forms.Label
    Friend WithEvents lblEtape As System.Windows.Forms.Label
    Friend WithEvents barre As System.Windows.Forms.ProgressBar
    Friend WithEvents lblCompteur As System.Windows.Forms.Label
    Friend WithEvents lblPatience As System.Windows.Forms.Label

End Class
