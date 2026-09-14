Option Strict On
Option Explicit On

Partial Class FrmFichierCoreBanking
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
        Me.panelEntete = New System.Windows.Forms.Panel()
        Me.lblTitre = New System.Windows.Forms.Label()
        Me.lblRecapitulatif = New System.Windows.Forms.Label()
        Me.dgvFichier = New System.Windows.Forms.DataGridView()
        Me.panelActions = New System.Windows.Forms.Panel()
        Me.btnExporter = New System.Windows.Forms.Button()
        Me.btnFermer = New System.Windows.Forms.Button()
        Me.lblEquilibre = New System.Windows.Forms.Label()
        Me.panelEntete.SuspendLayout()
        CType(Me.dgvFichier, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.panelActions.SuspendLayout()
        Me.SuspendLayout()
        '
        'panelEntete
        '
        Me.panelEntete.Controls.Add(Me.lblRecapitulatif)
        Me.panelEntete.Controls.Add(Me.lblTitre)
        Me.panelEntete.Dock = System.Windows.Forms.DockStyle.Top
        Me.panelEntete.Location = New System.Drawing.Point(0, 0)
        Me.panelEntete.Name = "panelEntete"
        Me.panelEntete.Size = New System.Drawing.Size(1100, 78)
        Me.panelEntete.TabIndex = 0
        '
        'lblTitre
        '
        Me.lblTitre.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Bold)
        Me.lblTitre.Location = New System.Drawing.Point(12, 10)
        Me.lblTitre.Name = "lblTitre"
        Me.lblTitre.Size = New System.Drawing.Size(1076, 24)
        Me.lblTitre.TabIndex = 0
        Me.lblTitre.Text = "Fichier destiné au core banking"
        '
        'lblRecapitulatif
        '
        Me.lblRecapitulatif.Location = New System.Drawing.Point(14, 38)
        Me.lblRecapitulatif.Name = "lblRecapitulatif"
        Me.lblRecapitulatif.Size = New System.Drawing.Size(1074, 36)
        Me.lblRecapitulatif.TabIndex = 1
        Me.lblRecapitulatif.Text = ""
        '
        'dgvFichier
        '
        Me.dgvFichier.AllowUserToAddRows = False
        Me.dgvFichier.AllowUserToDeleteRows = False
        Me.dgvFichier.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvFichier.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvFichier.Location = New System.Drawing.Point(0, 78)
        Me.dgvFichier.MultiSelect = False
        Me.dgvFichier.Name = "dgvFichier"
        Me.dgvFichier.ReadOnly = True
        Me.dgvFichier.RowHeadersWidth = 25
        Me.dgvFichier.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvFichier.Size = New System.Drawing.Size(1100, 474)
        Me.dgvFichier.TabIndex = 1
        '
        'panelActions
        '
        Me.panelActions.Controls.Add(Me.lblEquilibre)
        Me.panelActions.Controls.Add(Me.btnFermer)
        Me.panelActions.Controls.Add(Me.btnExporter)
        Me.panelActions.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.panelActions.Location = New System.Drawing.Point(0, 552)
        Me.panelActions.Name = "panelActions"
        Me.panelActions.Size = New System.Drawing.Size(1100, 48)
        Me.panelActions.TabIndex = 2
        '
        'btnExporter
        '
        Me.btnExporter.Location = New System.Drawing.Point(12, 8)
        Me.btnExporter.Name = "btnExporter"
        Me.btnExporter.Size = New System.Drawing.Size(250, 32)
        Me.btnExporter.TabIndex = 0
        Me.btnExporter.Text = "&Exporter vers Excel et ouvrir"
        Me.btnExporter.UseVisualStyleBackColor = True
        '
        'btnFermer
        '
        Me.btnFermer.Location = New System.Drawing.Point(270, 8)
        Me.btnFermer.Name = "btnFermer"
        Me.btnFermer.Size = New System.Drawing.Size(120, 32)
        Me.btnFermer.TabIndex = 1
        Me.btnFermer.Text = "&Fermer"
        Me.btnFermer.UseVisualStyleBackColor = True
        '
        'lblEquilibre
        '
        Me.lblEquilibre.Location = New System.Drawing.Point(400, 14)
        Me.lblEquilibre.Name = "lblEquilibre"
        Me.lblEquilibre.Size = New System.Drawing.Size(688, 22)
        Me.lblEquilibre.TabIndex = 2
        Me.lblEquilibre.Text = ""
        Me.lblEquilibre.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'FrmFichierCoreBanking
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1100, 600)
        Me.Controls.Add(Me.dgvFichier)
        Me.Controls.Add(Me.panelActions)
        Me.Controls.Add(Me.panelEntete)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        Me.MinimizeBox = False
        Me.Name = "FrmFichierCoreBanking"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Fichier core banking — consultation avant export"
        Me.panelEntete.ResumeLayout(False)
        CType(Me.dgvFichier, System.ComponentModel.ISupportInitialize).EndInit()
        Me.panelActions.ResumeLayout(False)
        Me.ResumeLayout(False)
    End Sub

    Friend WithEvents panelEntete As System.Windows.Forms.Panel
    Friend WithEvents lblTitre As System.Windows.Forms.Label
    Friend WithEvents lblRecapitulatif As System.Windows.Forms.Label
    Friend WithEvents dgvFichier As System.Windows.Forms.DataGridView
    Friend WithEvents panelActions As System.Windows.Forms.Panel
    Friend WithEvents btnExporter As System.Windows.Forms.Button
    Friend WithEvents btnFermer As System.Windows.Forms.Button
    Friend WithEvents lblEquilibre As System.Windows.Forms.Label

End Class
