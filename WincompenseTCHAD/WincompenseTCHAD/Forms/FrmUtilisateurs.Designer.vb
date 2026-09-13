Option Strict On
Option Explicit On

Partial Class FrmUtilisateurs
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
        Me.onglets = New System.Windows.Forms.TabControl()
        Me.pageComptes = New System.Windows.Forms.TabPage()
        Me.dgvUtilisateurs = New System.Windows.Forms.DataGridView()
        Me.panelActions = New System.Windows.Forms.Panel()
        Me.btnNouveau = New System.Windows.Forms.Button()
        Me.btnModifier = New System.Windows.Forms.Button()
        Me.btnReinitialiser = New System.Windows.Forms.Button()
        Me.btnDeverrouiller = New System.Windows.Forms.Button()
        Me.btnActualiser = New System.Windows.Forms.Button()
        Me.lblMessage = New System.Windows.Forms.Label()
        Me.pageJournal = New System.Windows.Forms.TabPage()
        Me.dgvConnexions = New System.Windows.Forms.DataGridView()
        Me.panelJournal = New System.Windows.Forms.Panel()
        Me.lblLignes = New System.Windows.Forms.Label()
        Me.cboLignes = New System.Windows.Forms.ComboBox()
        Me.btnActualiserJournal = New System.Windows.Forms.Button()
        Me.lblMessageJournal = New System.Windows.Forms.Label()
        Me.onglets.SuspendLayout()
        Me.pageComptes.SuspendLayout()
        CType(Me.dgvUtilisateurs, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.panelActions.SuspendLayout()
        Me.pageJournal.SuspendLayout()
        CType(Me.dgvConnexions, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.panelJournal.SuspendLayout()
        Me.SuspendLayout()
        '
        'onglets
        '
        Me.onglets.Controls.Add(Me.pageComptes)
        Me.onglets.Controls.Add(Me.pageJournal)
        Me.onglets.Dock = System.Windows.Forms.DockStyle.Fill
        Me.onglets.Location = New System.Drawing.Point(0, 0)
        Me.onglets.Name = "onglets"
        Me.onglets.SelectedIndex = 0
        Me.onglets.Size = New System.Drawing.Size(980, 600)
        Me.onglets.TabIndex = 0
        '
        'pageComptes
        '
        Me.pageComptes.Controls.Add(Me.dgvUtilisateurs)
        Me.pageComptes.Controls.Add(Me.panelActions)
        Me.pageComptes.Location = New System.Drawing.Point(4, 22)
        Me.pageComptes.Name = "pageComptes"
        Me.pageComptes.Padding = New System.Windows.Forms.Padding(6)
        Me.pageComptes.Size = New System.Drawing.Size(972, 574)
        Me.pageComptes.TabIndex = 0
        Me.pageComptes.Text = "Comptes utilisateurs"
        Me.pageComptes.UseVisualStyleBackColor = True
        '
        'dgvUtilisateurs
        '
        Me.dgvUtilisateurs.AllowUserToAddRows = False
        Me.dgvUtilisateurs.AllowUserToDeleteRows = False
        Me.dgvUtilisateurs.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvUtilisateurs.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvUtilisateurs.Location = New System.Drawing.Point(6, 6)
        Me.dgvUtilisateurs.MultiSelect = False
        Me.dgvUtilisateurs.Name = "dgvUtilisateurs"
        Me.dgvUtilisateurs.ReadOnly = True
        Me.dgvUtilisateurs.RowHeadersWidth = 25
        Me.dgvUtilisateurs.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvUtilisateurs.Size = New System.Drawing.Size(960, 494)
        Me.dgvUtilisateurs.TabIndex = 0
        '
        'panelActions
        '
        Me.panelActions.Controls.Add(Me.lblMessage)
        Me.panelActions.Controls.Add(Me.btnActualiser)
        Me.panelActions.Controls.Add(Me.btnDeverrouiller)
        Me.panelActions.Controls.Add(Me.btnReinitialiser)
        Me.panelActions.Controls.Add(Me.btnModifier)
        Me.panelActions.Controls.Add(Me.btnNouveau)
        Me.panelActions.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.panelActions.Location = New System.Drawing.Point(6, 500)
        Me.panelActions.Name = "panelActions"
        Me.panelActions.Size = New System.Drawing.Size(960, 68)
        Me.panelActions.TabIndex = 1
        '
        'btnNouveau
        '
        Me.btnNouveau.Location = New System.Drawing.Point(4, 6)
        Me.btnNouveau.Name = "btnNouveau"
        Me.btnNouveau.Size = New System.Drawing.Size(150, 32)
        Me.btnNouveau.TabIndex = 0
        Me.btnNouveau.Text = "&Nouvel utilisateur"
        Me.btnNouveau.UseVisualStyleBackColor = True
        '
        'btnModifier
        '
        Me.btnModifier.Location = New System.Drawing.Point(162, 6)
        Me.btnModifier.Name = "btnModifier"
        Me.btnModifier.Size = New System.Drawing.Size(150, 32)
        Me.btnModifier.TabIndex = 1
        Me.btnModifier.Text = "&Modifier"
        Me.btnModifier.UseVisualStyleBackColor = True
        '
        'btnReinitialiser
        '
        Me.btnReinitialiser.Location = New System.Drawing.Point(320, 6)
        Me.btnReinitialiser.Name = "btnReinitialiser"
        Me.btnReinitialiser.Size = New System.Drawing.Size(202, 32)
        Me.btnReinitialiser.TabIndex = 2
        Me.btnReinitialiser.Text = "&Réinitialiser le mot de passe"
        Me.btnReinitialiser.UseVisualStyleBackColor = True
        '
        'btnDeverrouiller
        '
        Me.btnDeverrouiller.Location = New System.Drawing.Point(530, 6)
        Me.btnDeverrouiller.Name = "btnDeverrouiller"
        Me.btnDeverrouiller.Size = New System.Drawing.Size(150, 32)
        Me.btnDeverrouiller.TabIndex = 3
        Me.btnDeverrouiller.Text = "&Déverrouiller"
        Me.btnDeverrouiller.UseVisualStyleBackColor = True
        '
        'btnActualiser
        '
        Me.btnActualiser.Location = New System.Drawing.Point(688, 6)
        Me.btnActualiser.Name = "btnActualiser"
        Me.btnActualiser.Size = New System.Drawing.Size(150, 32)
        Me.btnActualiser.TabIndex = 4
        Me.btnActualiser.Text = "&Actualiser"
        Me.btnActualiser.UseVisualStyleBackColor = True
        '
        'lblMessage
        '
        Me.lblMessage.ForeColor = System.Drawing.Color.Firebrick
        Me.lblMessage.Location = New System.Drawing.Point(4, 42)
        Me.lblMessage.Name = "lblMessage"
        Me.lblMessage.Size = New System.Drawing.Size(950, 22)
        Me.lblMessage.TabIndex = 5
        Me.lblMessage.Text = ""
        '
        'pageJournal
        '
        Me.pageJournal.Controls.Add(Me.dgvConnexions)
        Me.pageJournal.Controls.Add(Me.panelJournal)
        Me.pageJournal.Location = New System.Drawing.Point(4, 22)
        Me.pageJournal.Name = "pageJournal"
        Me.pageJournal.Padding = New System.Windows.Forms.Padding(6)
        Me.pageJournal.Size = New System.Drawing.Size(972, 574)
        Me.pageJournal.TabIndex = 1
        Me.pageJournal.Text = "Journal des connexions"
        Me.pageJournal.UseVisualStyleBackColor = True
        '
        'dgvConnexions
        '
        Me.dgvConnexions.AllowUserToAddRows = False
        Me.dgvConnexions.AllowUserToDeleteRows = False
        Me.dgvConnexions.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvConnexions.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvConnexions.Location = New System.Drawing.Point(6, 6)
        Me.dgvConnexions.MultiSelect = False
        Me.dgvConnexions.Name = "dgvConnexions"
        Me.dgvConnexions.ReadOnly = True
        Me.dgvConnexions.RowHeadersWidth = 25
        Me.dgvConnexions.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvConnexions.Size = New System.Drawing.Size(960, 512)
        Me.dgvConnexions.TabIndex = 0
        '
        'panelJournal
        '
        Me.panelJournal.Controls.Add(Me.lblMessageJournal)
        Me.panelJournal.Controls.Add(Me.btnActualiserJournal)
        Me.panelJournal.Controls.Add(Me.cboLignes)
        Me.panelJournal.Controls.Add(Me.lblLignes)
        Me.panelJournal.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.panelJournal.Location = New System.Drawing.Point(6, 518)
        Me.panelJournal.Name = "panelJournal"
        Me.panelJournal.Size = New System.Drawing.Size(960, 50)
        Me.panelJournal.TabIndex = 1
        '
        'lblLignes
        '
        Me.lblLignes.Location = New System.Drawing.Point(4, 12)
        Me.lblLignes.Name = "lblLignes"
        Me.lblLignes.Size = New System.Drawing.Size(180, 24)
        Me.lblLignes.TabIndex = 0
        Me.lblLignes.Text = "Dernières tentatives affichées"
        Me.lblLignes.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'cboLignes
        '
        Me.cboLignes.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboLignes.FormattingEnabled = True
        Me.cboLignes.Location = New System.Drawing.Point(190, 10)
        Me.cboLignes.Name = "cboLignes"
        Me.cboLignes.Size = New System.Drawing.Size(110, 24)
        Me.cboLignes.TabIndex = 1
        '
        'btnActualiserJournal
        '
        Me.btnActualiserJournal.Location = New System.Drawing.Point(312, 8)
        Me.btnActualiserJournal.Name = "btnActualiserJournal"
        Me.btnActualiserJournal.Size = New System.Drawing.Size(150, 32)
        Me.btnActualiserJournal.TabIndex = 2
        Me.btnActualiserJournal.Text = "A&ctualiser le journal"
        Me.btnActualiserJournal.UseVisualStyleBackColor = True
        '
        'lblMessageJournal
        '
        Me.lblMessageJournal.ForeColor = System.Drawing.Color.Firebrick
        Me.lblMessageJournal.Location = New System.Drawing.Point(474, 12)
        Me.lblMessageJournal.Name = "lblMessageJournal"
        Me.lblMessageJournal.Size = New System.Drawing.Size(480, 24)
        Me.lblMessageJournal.TabIndex = 3
        Me.lblMessageJournal.Text = ""
        Me.lblMessageJournal.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'FrmUtilisateurs
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(980, 600)
        Me.Controls.Add(Me.onglets)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        Me.Name = "FrmUtilisateurs"
        Me.Text = "Utilisateurs et connexions"
        Me.onglets.ResumeLayout(False)
        Me.pageComptes.ResumeLayout(False)
        CType(Me.dgvUtilisateurs, System.ComponentModel.ISupportInitialize).EndInit()
        Me.panelActions.ResumeLayout(False)
        Me.pageJournal.ResumeLayout(False)
        CType(Me.dgvConnexions, System.ComponentModel.ISupportInitialize).EndInit()
        Me.panelJournal.ResumeLayout(False)
        Me.ResumeLayout(False)
    End Sub

    Friend WithEvents onglets As System.Windows.Forms.TabControl
    Friend WithEvents pageComptes As System.Windows.Forms.TabPage
    Friend WithEvents dgvUtilisateurs As System.Windows.Forms.DataGridView
    Friend WithEvents panelActions As System.Windows.Forms.Panel
    Friend WithEvents btnNouveau As System.Windows.Forms.Button
    Friend WithEvents btnModifier As System.Windows.Forms.Button
    Friend WithEvents btnReinitialiser As System.Windows.Forms.Button
    Friend WithEvents btnDeverrouiller As System.Windows.Forms.Button
    Friend WithEvents btnActualiser As System.Windows.Forms.Button
    Friend WithEvents lblMessage As System.Windows.Forms.Label
    Friend WithEvents pageJournal As System.Windows.Forms.TabPage
    Friend WithEvents dgvConnexions As System.Windows.Forms.DataGridView
    Friend WithEvents panelJournal As System.Windows.Forms.Panel
    Friend WithEvents lblLignes As System.Windows.Forms.Label
    Friend WithEvents cboLignes As System.Windows.Forms.ComboBox
    Friend WithEvents btnActualiserJournal As System.Windows.Forms.Button
    Friend WithEvents lblMessageJournal As System.Windows.Forms.Label

End Class
