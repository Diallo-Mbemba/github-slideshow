Option Strict On
Option Explicit On

Partial Class FrmPrincipal
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
        Me.menuPrincipal = New System.Windows.Forms.MenuStrip()
        Me.mnuCompensation = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuTraitement = New System.Windows.Forms.ToolStripMenuItem()
        Me.SEP1 = New System.Windows.Forms.ToolStripSeparator()
        Me.mnuRapport = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuParametrage = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuSousAgents = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuAgences = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuGroupes = New System.Windows.Forms.ToolStripMenuItem()
        Me.SEP5 = New System.Windows.Forms.ToolStripSeparator()
        Me.mnuDemandes = New System.Windows.Forms.ToolStripMenuItem()
        Me.SEP2 = New System.Windows.Forms.ToolStripSeparator()
        Me.mnuComptes = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuSecurite = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuMonMotDePasse = New System.Windows.Forms.ToolStripMenuItem()
        Me.SEP4 = New System.Windows.Forms.ToolStripSeparator()
        Me.mnuUtilisateurs = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuFenetres = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuCascade = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuMosaiqueH = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuMosaiqueV = New System.Windows.Forms.ToolStripMenuItem()
        Me.SEP3 = New System.Windows.Forms.ToolStripSeparator()
        Me.mnuFermerTout = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuQuitter = New System.Windows.Forms.ToolStripMenuItem()
        Me.barreEtat = New System.Windows.Forms.StatusStrip()
        Me.tsslUtilisateur = New System.Windows.Forms.ToolStripStatusLabel()
        Me.tsslBase = New System.Windows.Forms.ToolStripStatusLabel()
        Me.menuPrincipal.SuspendLayout()
        Me.barreEtat.SuspendLayout()
        Me.SuspendLayout()
        '
        'menuPrincipal
        '
        Me.menuPrincipal.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuCompensation, Me.mnuParametrage, Me.mnuSecurite, Me.mnuFenetres, Me.mnuQuitter})
        Me.menuPrincipal.Location = New System.Drawing.Point(0, 0)
        Me.menuPrincipal.MdiWindowListItem = Me.mnuFenetres
        Me.menuPrincipal.Name = "menuPrincipal"
        Me.menuPrincipal.Size = New System.Drawing.Size(1200, 24)
        Me.menuPrincipal.TabIndex = 0
        '
        'mnuCompensation
        '
        Me.mnuCompensation.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuTraitement, Me.SEP1, Me.mnuRapport})
        Me.mnuCompensation.Name = "mnuCompensation"
        Me.mnuCompensation.Size = New System.Drawing.Size(100, 20)
        Me.mnuCompensation.Text = "&Compensation"
        '
        'mnuTraitement
        '
        Me.mnuTraitement.Name = "mnuTraitement"
        Me.mnuTraitement.Size = New System.Drawing.Size(240, 22)
        Me.mnuTraitement.Text = "&Traitement de la compense"
        '
        'SEP1
        '
        Me.SEP1.Name = "SEP1"
        Me.SEP1.Size = New System.Drawing.Size(220, 6)
        '
        'mnuRapport
        '
        Me.mnuRapport.Name = "mnuRapport"
        Me.mnuRapport.Size = New System.Drawing.Size(240, 22)
        Me.mnuRapport.Text = "&Rapport d'activité"
        '
        'mnuParametrage
        '
        Me.mnuParametrage.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuSousAgents, Me.mnuAgences, Me.mnuGroupes, Me.SEP5, Me.mnuDemandes, Me.SEP2, Me.mnuComptes})
        Me.mnuParametrage.Name = "mnuParametrage"
        Me.mnuParametrage.Size = New System.Drawing.Size(100, 20)
        Me.mnuParametrage.Text = "&Paramétrage"
        '
        'mnuSousAgents
        '
        Me.mnuSousAgents.Name = "mnuSousAgents"
        Me.mnuSousAgents.Size = New System.Drawing.Size(240, 22)
        Me.mnuSousAgents.Text = "&Sous-agents"
        '
        'mnuAgences
        '
        Me.mnuAgences.Name = "mnuAgences"
        Me.mnuAgences.Size = New System.Drawing.Size(240, 22)
        Me.mnuAgences.Text = "&Agences propres"
        '
        'mnuGroupes
        '
        Me.mnuGroupes.Name = "mnuGroupes"
        Me.mnuGroupes.Size = New System.Drawing.Size(240, 22)
        Me.mnuGroupes.Text = "&Groupes statistiques"
        '
        'SEP5
        '
        Me.SEP5.Name = "SEP5"
        Me.SEP5.Size = New System.Drawing.Size(220, 6)
        '
        'mnuDemandes
        '
        Me.mnuDemandes.Name = "mnuDemandes"
        Me.mnuDemandes.Size = New System.Drawing.Size(240, 22)
        Me.mnuDemandes.Text = "&Autorisations du référentiel"
        '
        'SEP2
        '
        Me.SEP2.Name = "SEP2"
        Me.SEP2.Size = New System.Drawing.Size(220, 6)
        '
        'mnuComptes
        '
        Me.mnuComptes.Name = "mnuComptes"
        Me.mnuComptes.Size = New System.Drawing.Size(240, 22)
        Me.mnuComptes.Text = "&Comptes systèmes"
        '
        'mnuSecurite
        '
        Me.mnuSecurite.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuMonMotDePasse, Me.SEP4, Me.mnuUtilisateurs})
        Me.mnuSecurite.Name = "mnuSecurite"
        Me.mnuSecurite.Size = New System.Drawing.Size(100, 20)
        Me.mnuSecurite.Text = "&Sécurité"
        '
        'mnuMonMotDePasse
        '
        Me.mnuMonMotDePasse.Name = "mnuMonMotDePasse"
        Me.mnuMonMotDePasse.Size = New System.Drawing.Size(260, 22)
        Me.mnuMonMotDePasse.Text = "&Mon mot de passe..."
        '
        'SEP4
        '
        Me.SEP4.Name = "SEP4"
        Me.SEP4.Size = New System.Drawing.Size(240, 6)
        '
        'mnuUtilisateurs
        '
        Me.mnuUtilisateurs.Name = "mnuUtilisateurs"
        Me.mnuUtilisateurs.Size = New System.Drawing.Size(260, 22)
        Me.mnuUtilisateurs.Text = "&Utilisateurs et connexions"
        '
        'mnuFenetres
        '
        Me.mnuFenetres.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuCascade, Me.mnuMosaiqueH, Me.mnuMosaiqueV, Me.SEP3, Me.mnuFermerTout})
        Me.mnuFenetres.Name = "mnuFenetres"
        Me.mnuFenetres.Size = New System.Drawing.Size(100, 20)
        Me.mnuFenetres.Text = "&Fenêtres"
        '
        'mnuCascade
        '
        Me.mnuCascade.Name = "mnuCascade"
        Me.mnuCascade.Size = New System.Drawing.Size(240, 22)
        Me.mnuCascade.Text = "&Cascade"
        '
        'mnuMosaiqueH
        '
        Me.mnuMosaiqueH.Name = "mnuMosaiqueH"
        Me.mnuMosaiqueH.Size = New System.Drawing.Size(240, 22)
        Me.mnuMosaiqueH.Text = "Mosaïque &horizontale"
        '
        'mnuMosaiqueV
        '
        Me.mnuMosaiqueV.Name = "mnuMosaiqueV"
        Me.mnuMosaiqueV.Size = New System.Drawing.Size(240, 22)
        Me.mnuMosaiqueV.Text = "Mosaïque &verticale"
        '
        'SEP3
        '
        Me.SEP3.Name = "SEP3"
        Me.SEP3.Size = New System.Drawing.Size(220, 6)
        '
        'mnuFermerTout
        '
        Me.mnuFermerTout.Name = "mnuFermerTout"
        Me.mnuFermerTout.Size = New System.Drawing.Size(240, 22)
        Me.mnuFermerTout.Text = "&Fermer toutes les fenêtres"
        '
        'mnuQuitter
        '
        Me.mnuQuitter.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.mnuQuitter.Name = "mnuQuitter"
        Me.mnuQuitter.Size = New System.Drawing.Size(70, 20)
        Me.mnuQuitter.Text = "&Quitter"
        '
        'barreEtat
        '
        Me.barreEtat.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsslUtilisateur, Me.tsslBase})
        Me.barreEtat.Location = New System.Drawing.Point(0, 678)
        Me.barreEtat.Name = "barreEtat"
        Me.barreEtat.Size = New System.Drawing.Size(1200, 22)
        Me.barreEtat.TabIndex = 1
        '
        'tsslUtilisateur
        '
        Me.tsslUtilisateur.Name = "tsslUtilisateur"
        Me.tsslUtilisateur.Size = New System.Drawing.Size(600, 17)
        Me.tsslUtilisateur.Text = ""
        Me.tsslUtilisateur.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'tsslBase
        '
        Me.tsslBase.Name = "tsslBase"
        Me.tsslBase.Size = New System.Drawing.Size(585, 17)
        Me.tsslBase.Spring = True
        Me.tsslBase.Text = ""
        Me.tsslBase.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'FrmPrincipal
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1200, 700)
        Me.Controls.Add(Me.barreEtat)
        Me.Controls.Add(Me.menuPrincipal)
        Me.IsMdiContainer = True
        Me.MainMenuStrip = Me.menuPrincipal
        Me.Name = "FrmPrincipal"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Wincompense TCHAD — Compensation Western Union"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.menuPrincipal.ResumeLayout(False)
        Me.menuPrincipal.PerformLayout()
        Me.barreEtat.ResumeLayout(False)
        Me.barreEtat.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()
    End Sub

    Friend WithEvents menuPrincipal As System.Windows.Forms.MenuStrip
    Friend WithEvents mnuCompensation As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuTraitement As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SEP1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents mnuRapport As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuParametrage As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuSousAgents As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuAgences As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuGroupes As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SEP5 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents mnuDemandes As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SEP2 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents mnuComptes As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuSecurite As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuMonMotDePasse As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SEP4 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents mnuUtilisateurs As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuFenetres As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuCascade As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuMosaiqueH As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuMosaiqueV As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SEP3 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents mnuFermerTout As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuQuitter As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents barreEtat As System.Windows.Forms.StatusStrip
    Friend WithEvents tsslUtilisateur As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tsslBase As System.Windows.Forms.ToolStripStatusLabel

End Class
