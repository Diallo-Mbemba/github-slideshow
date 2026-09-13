Option Strict On
Option Explicit On

Partial Class FrmDemandes
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
        Me.pageAttente = New System.Windows.Forms.TabPage()
        Me.separateur = New System.Windows.Forms.SplitContainer()
        Me.dgvAttente = New System.Windows.Forms.DataGridView()
        Me.dgvDetail = New System.Windows.Forms.DataGridView()
        Me.panelActions = New System.Windows.Forms.Panel()
        Me.btnAutoriser = New System.Windows.Forms.Button()
        Me.btnRejeter = New System.Windows.Forms.Button()
        Me.btnActualiser = New System.Windows.Forms.Button()
        Me.lblMessage = New System.Windows.Forms.Label()
        Me.pageHistorique = New System.Windows.Forms.TabPage()
        Me.dgvHistorique = New System.Windows.Forms.DataGridView()
        Me.panelHistorique = New System.Windows.Forms.Panel()
        Me.btnActualiserHistorique = New System.Windows.Forms.Button()
        Me.lblMessageHistorique = New System.Windows.Forms.Label()
        Me.onglets.SuspendLayout()
        Me.pageAttente.SuspendLayout()
        CType(Me.separateur, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.separateur.Panel1.SuspendLayout()
        Me.separateur.Panel2.SuspendLayout()
        Me.separateur.SuspendLayout()
        CType(Me.dgvAttente, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvDetail, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.panelActions.SuspendLayout()
        Me.pageHistorique.SuspendLayout()
        CType(Me.dgvHistorique, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.panelHistorique.SuspendLayout()
        Me.SuspendLayout()
        '
        'onglets
        '
        Me.onglets.Controls.Add(Me.pageAttente)
        Me.onglets.Controls.Add(Me.pageHistorique)
        Me.onglets.Dock = System.Windows.Forms.DockStyle.Fill
        Me.onglets.Location = New System.Drawing.Point(0, 0)
        Me.onglets.Name = "onglets"
        Me.onglets.SelectedIndex = 0
        Me.onglets.Size = New System.Drawing.Size(1000, 620)
        Me.onglets.TabIndex = 0
        '
        'pageAttente
        '
        Me.pageAttente.Controls.Add(Me.separateur)
        Me.pageAttente.Controls.Add(Me.panelActions)
        Me.pageAttente.Location = New System.Drawing.Point(4, 22)
        Me.pageAttente.Name = "pageAttente"
        Me.pageAttente.Padding = New System.Windows.Forms.Padding(6)
        Me.pageAttente.Size = New System.Drawing.Size(992, 594)
        Me.pageAttente.TabIndex = 0
        Me.pageAttente.Text = "Demandes en attente"
        Me.pageAttente.UseVisualStyleBackColor = True
        '
        'separateur
        '
        Me.separateur.Dock = System.Windows.Forms.DockStyle.Fill
        Me.separateur.Location = New System.Drawing.Point(6, 6)
        Me.separateur.Name = "separateur"
        Me.separateur.Orientation = System.Windows.Forms.Orientation.Horizontal
        Me.separateur.Panel1.Controls.Add(Me.dgvAttente)
        Me.separateur.Panel2.Controls.Add(Me.dgvDetail)
        Me.separateur.Size = New System.Drawing.Size(980, 514)
        Me.separateur.SplitterDistance = 260
        Me.separateur.TabIndex = 0
        '
        'dgvAttente
        '
        Me.dgvAttente.AllowUserToAddRows = False
        Me.dgvAttente.AllowUserToDeleteRows = False
        Me.dgvAttente.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvAttente.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvAttente.Location = New System.Drawing.Point(0, 0)
        Me.dgvAttente.MultiSelect = False
        Me.dgvAttente.Name = "dgvAttente"
        Me.dgvAttente.ReadOnly = True
        Me.dgvAttente.RowHeadersWidth = 25
        Me.dgvAttente.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvAttente.Size = New System.Drawing.Size(980, 260)
        Me.dgvAttente.TabIndex = 0
        '
        'dgvDetail
        '
        Me.dgvDetail.AllowUserToAddRows = False
        Me.dgvDetail.AllowUserToDeleteRows = False
        Me.dgvDetail.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvDetail.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvDetail.Location = New System.Drawing.Point(0, 0)
        Me.dgvDetail.MultiSelect = False
        Me.dgvDetail.Name = "dgvDetail"
        Me.dgvDetail.ReadOnly = True
        Me.dgvDetail.RowHeadersVisible = False
        Me.dgvDetail.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvDetail.Size = New System.Drawing.Size(980, 250)
        Me.dgvDetail.TabIndex = 0
        '
        'panelActions
        '
        Me.panelActions.Controls.Add(Me.lblMessage)
        Me.panelActions.Controls.Add(Me.btnActualiser)
        Me.panelActions.Controls.Add(Me.btnRejeter)
        Me.panelActions.Controls.Add(Me.btnAutoriser)
        Me.panelActions.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.panelActions.Location = New System.Drawing.Point(6, 520)
        Me.panelActions.Name = "panelActions"
        Me.panelActions.Size = New System.Drawing.Size(980, 68)
        Me.panelActions.TabIndex = 1
        '
        'btnAutoriser
        '
        Me.btnAutoriser.Location = New System.Drawing.Point(4, 6)
        Me.btnAutoriser.Name = "btnAutoriser"
        Me.btnAutoriser.Size = New System.Drawing.Size(170, 32)
        Me.btnAutoriser.TabIndex = 0
        Me.btnAutoriser.Text = "&Autoriser"
        Me.btnAutoriser.UseVisualStyleBackColor = True
        '
        'btnRejeter
        '
        Me.btnRejeter.Location = New System.Drawing.Point(182, 6)
        Me.btnRejeter.Name = "btnRejeter"
        Me.btnRejeter.Size = New System.Drawing.Size(170, 32)
        Me.btnRejeter.TabIndex = 1
        Me.btnRejeter.Text = "&Rejeter..."
        Me.btnRejeter.UseVisualStyleBackColor = True
        '
        'btnActualiser
        '
        Me.btnActualiser.Location = New System.Drawing.Point(360, 6)
        Me.btnActualiser.Name = "btnActualiser"
        Me.btnActualiser.Size = New System.Drawing.Size(150, 32)
        Me.btnActualiser.TabIndex = 2
        Me.btnActualiser.Text = "A&ctualiser"
        Me.btnActualiser.UseVisualStyleBackColor = True
        '
        'lblMessage
        '
        Me.lblMessage.ForeColor = System.Drawing.Color.Firebrick
        Me.lblMessage.Location = New System.Drawing.Point(4, 42)
        Me.lblMessage.Name = "lblMessage"
        Me.lblMessage.Size = New System.Drawing.Size(970, 22)
        Me.lblMessage.TabIndex = 3
        Me.lblMessage.Text = ""
        '
        'pageHistorique
        '
        Me.pageHistorique.Controls.Add(Me.dgvHistorique)
        Me.pageHistorique.Controls.Add(Me.panelHistorique)
        Me.pageHistorique.Location = New System.Drawing.Point(4, 22)
        Me.pageHistorique.Name = "pageHistorique"
        Me.pageHistorique.Padding = New System.Windows.Forms.Padding(6)
        Me.pageHistorique.Size = New System.Drawing.Size(992, 594)
        Me.pageHistorique.TabIndex = 1
        Me.pageHistorique.Text = "Demandes décidées"
        Me.pageHistorique.UseVisualStyleBackColor = True
        '
        'dgvHistorique
        '
        Me.dgvHistorique.AllowUserToAddRows = False
        Me.dgvHistorique.AllowUserToDeleteRows = False
        Me.dgvHistorique.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvHistorique.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvHistorique.Location = New System.Drawing.Point(6, 6)
        Me.dgvHistorique.MultiSelect = False
        Me.dgvHistorique.Name = "dgvHistorique"
        Me.dgvHistorique.ReadOnly = True
        Me.dgvHistorique.RowHeadersWidth = 25
        Me.dgvHistorique.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvHistorique.Size = New System.Drawing.Size(980, 532)
        Me.dgvHistorique.TabIndex = 0
        '
        'panelHistorique
        '
        Me.panelHistorique.Controls.Add(Me.lblMessageHistorique)
        Me.panelHistorique.Controls.Add(Me.btnActualiserHistorique)
        Me.panelHistorique.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.panelHistorique.Location = New System.Drawing.Point(6, 538)
        Me.panelHistorique.Name = "panelHistorique"
        Me.panelHistorique.Size = New System.Drawing.Size(980, 50)
        Me.panelHistorique.TabIndex = 1
        '
        'btnActualiserHistorique
        '
        Me.btnActualiserHistorique.Location = New System.Drawing.Point(4, 8)
        Me.btnActualiserHistorique.Name = "btnActualiserHistorique"
        Me.btnActualiserHistorique.Size = New System.Drawing.Size(150, 32)
        Me.btnActualiserHistorique.TabIndex = 0
        Me.btnActualiserHistorique.Text = "&Actualiser"
        Me.btnActualiserHistorique.UseVisualStyleBackColor = True
        '
        'lblMessageHistorique
        '
        Me.lblMessageHistorique.ForeColor = System.Drawing.Color.Firebrick
        Me.lblMessageHistorique.Location = New System.Drawing.Point(166, 12)
        Me.lblMessageHistorique.Name = "lblMessageHistorique"
        Me.lblMessageHistorique.Size = New System.Drawing.Size(808, 24)
        Me.lblMessageHistorique.TabIndex = 1
        Me.lblMessageHistorique.Text = ""
        Me.lblMessageHistorique.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'FrmDemandes
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1000, 620)
        Me.Controls.Add(Me.onglets)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        Me.Name = "FrmDemandes"
        Me.Text = "Autorisations du référentiel"
        Me.onglets.ResumeLayout(False)
        Me.pageAttente.ResumeLayout(False)
        Me.separateur.Panel1.ResumeLayout(False)
        Me.separateur.Panel2.ResumeLayout(False)
        CType(Me.separateur, System.ComponentModel.ISupportInitialize).EndInit()
        Me.separateur.ResumeLayout(False)
        CType(Me.dgvAttente, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvDetail, System.ComponentModel.ISupportInitialize).EndInit()
        Me.panelActions.ResumeLayout(False)
        Me.pageHistorique.ResumeLayout(False)
        CType(Me.dgvHistorique, System.ComponentModel.ISupportInitialize).EndInit()
        Me.panelHistorique.ResumeLayout(False)
        Me.ResumeLayout(False)
    End Sub

    Friend WithEvents onglets As System.Windows.Forms.TabControl
    Friend WithEvents pageAttente As System.Windows.Forms.TabPage
    Friend WithEvents separateur As System.Windows.Forms.SplitContainer
    Friend WithEvents dgvAttente As System.Windows.Forms.DataGridView
    Friend WithEvents dgvDetail As System.Windows.Forms.DataGridView
    Friend WithEvents panelActions As System.Windows.Forms.Panel
    Friend WithEvents btnAutoriser As System.Windows.Forms.Button
    Friend WithEvents btnRejeter As System.Windows.Forms.Button
    Friend WithEvents btnActualiser As System.Windows.Forms.Button
    Friend WithEvents lblMessage As System.Windows.Forms.Label
    Friend WithEvents pageHistorique As System.Windows.Forms.TabPage
    Friend WithEvents dgvHistorique As System.Windows.Forms.DataGridView
    Friend WithEvents panelHistorique As System.Windows.Forms.Panel
    Friend WithEvents btnActualiserHistorique As System.Windows.Forms.Button
    Friend WithEvents lblMessageHistorique As System.Windows.Forms.Label

End Class
