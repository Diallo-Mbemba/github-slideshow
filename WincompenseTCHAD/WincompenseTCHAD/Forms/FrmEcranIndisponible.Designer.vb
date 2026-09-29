Option Strict On
Option Explicit On

Partial Class FrmEcranIndisponible
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
        Me.lblExplication = New System.Windows.Forms.Label()
        Me.lblCeQuiMarche = New System.Windows.Forms.Label()
        Me.btnFermer = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'lblTitre
        '
        Me.lblTitre.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblTitre.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Bold)
        Me.lblTitre.Location = New System.Drawing.Point(20, 22)
        Me.lblTitre.Name = "lblTitre"
        Me.lblTitre.Size = New System.Drawing.Size(516, 46)
        Me.lblTitre.TabIndex = 0
        Me.lblTitre.Text = "Écran non disponible"
        '
        'lblExplication
        '
        Me.lblExplication.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblExplication.Location = New System.Drawing.Point(22, 78)
        Me.lblExplication.Name = "lblExplication"
        Me.lblExplication.Size = New System.Drawing.Size(516, 60)
        Me.lblExplication.TabIndex = 1
        Me.lblExplication.Text = ""
        '
        'lblCeQuiMarche
        '
        Me.lblCeQuiMarche.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblCeQuiMarche.ForeColor = System.Drawing.SystemColors.GrayText
        Me.lblCeQuiMarche.Location = New System.Drawing.Point(22, 148)
        Me.lblCeQuiMarche.Name = "lblCeQuiMarche"
        Me.lblCeQuiMarche.Size = New System.Drawing.Size(516, 56)
        Me.lblCeQuiMarche.TabIndex = 2
        Me.lblCeQuiMarche.Text = "Le menu « Changer de produit » ramène au choix du produit, sans quitter l'applicat" & _
            "ion. La sécurité, les utilisateurs et la connexion à la base restent accessibles :" & _
            " ils ne dépendent pas du produit."
        '
        'btnFermer
        '
        Me.btnFermer.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFermer.Location = New System.Drawing.Point(438, 218)
        Me.btnFermer.Name = "btnFermer"
        Me.btnFermer.Size = New System.Drawing.Size(100, 30)
        Me.btnFermer.TabIndex = 3
        Me.btnFermer.Text = "&Fermer"
        Me.btnFermer.UseVisualStyleBackColor = True
        '
        'FrmEcranIndisponible
        '
        Me.AcceptButton = Me.btnFermer
        Me.CancelButton = Me.btnFermer
        Me.ClientSize = New System.Drawing.Size(560, 266)
        Me.Controls.Add(Me.btnFermer)
        Me.Controls.Add(Me.lblCeQuiMarche)
        Me.Controls.Add(Me.lblExplication)
        Me.Controls.Add(Me.lblTitre)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(480, 260)
        Me.Name = "FrmEcranIndisponible"
        Me.Text = "Écran non disponible"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents lblTitre As System.Windows.Forms.Label
    Friend WithEvents lblExplication As System.Windows.Forms.Label
    Friend WithEvents lblCeQuiMarche As System.Windows.Forms.Label
    Friend WithEvents btnFermer As System.Windows.Forms.Button

End Class
