Option Strict On
Option Explicit On

Partial Class FrmComptesSysteme
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
        Me.lblCompteCourant = New System.Windows.Forms.Label()
        Me.cboCompteCourant = New System.Windows.Forms.ComboBox()
        Me.lblInterBancaire = New System.Windows.Forms.Label()
        Me.cboInterBancaire = New System.Windows.Forms.ComboBox()
        Me.lblCommissionTransfert = New System.Windows.Forms.Label()
        Me.cboCommissionTransfert = New System.Windows.Forms.ComboBox()
        Me.lblCommissionEnvoi = New System.Windows.Forms.Label()
        Me.cboCommissionEnvoi = New System.Windows.Forms.ComboBox()
        Me.lblCommissionPaiement = New System.Windows.Forms.Label()
        Me.cboCommissionPaiement = New System.Windows.Forms.ComboBox()
        Me.lblImpotsTaxeEnvoi = New System.Windows.Forms.Label()
        Me.cboImpotsTaxeEnvoi = New System.Windows.Forms.ComboBox()
        Me.lblTVA = New System.Windows.Forms.Label()
        Me.cboTVA = New System.Windows.Forms.ComboBox()
        Me.lblTTAEnvoi = New System.Windows.Forms.Label()
        Me.cboTTAEnvoi = New System.Windows.Forms.ComboBox()
        Me.lblTTAReception = New System.Windows.Forms.Label()
        Me.cboTTAReception = New System.Windows.Forms.ComboBox()
        Me.btnEnregistrer = New System.Windows.Forms.Button()
        Me.btnFermer = New System.Windows.Forms.Button()
        Me.lblInfo = New System.Windows.Forms.Label()
        Me.panelTitre.SuspendLayout()
        Me.SuspendLayout()
        '
        'panelTitre
        '
        Me.panelTitre.BackColor = System.Drawing.Color.Black
        Me.panelTitre.Controls.Add(Me.lblTitre)
        Me.panelTitre.Location = New System.Drawing.Point(12, 12)
        Me.panelTitre.Name = "panelTitre"
        Me.panelTitre.Size = New System.Drawing.Size(410, 62)
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
        Me.lblTitre.Size = New System.Drawing.Size(410, 62)
        Me.lblTitre.TabIndex = 0
        Me.lblTitre.Text = "Comptes Systemes WU"
        Me.lblTitre.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblCompteCourant
        '
        Me.lblCompteCourant.Location = New System.Drawing.Point(18, 99)
        Me.lblCompteCourant.Name = "lblCompteCourant"
        Me.lblCompteCourant.Size = New System.Drawing.Size(256, 30)
        Me.lblCompteCourant.TabIndex = 1
        Me.lblCompteCourant.Text = "COMPTE COURANT WESTERN UNION ETD"
        Me.lblCompteCourant.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'cboCompteCourant
        '
        Me.cboCompteCourant.FormattingEnabled = True
        Me.cboCompteCourant.Location = New System.Drawing.Point(280, 95)
        Me.cboCompteCourant.MaxLength = 255
        Me.cboCompteCourant.Name = "cboCompteCourant"
        Me.cboCompteCourant.Size = New System.Drawing.Size(142, 24)
        Me.cboCompteCourant.TabIndex = 2
        '
        'lblInterBancaire
        '
        Me.lblInterBancaire.Location = New System.Drawing.Point(18, 137)
        Me.lblInterBancaire.Name = "lblInterBancaire"
        Me.lblInterBancaire.Size = New System.Drawing.Size(256, 30)
        Me.lblInterBancaire.TabIndex = 3
        Me.lblInterBancaire.Text = "COMPTE INTER BANCAIRE"
        Me.lblInterBancaire.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'cboInterBancaire
        '
        Me.cboInterBancaire.FormattingEnabled = True
        Me.cboInterBancaire.Location = New System.Drawing.Point(280, 133)
        Me.cboInterBancaire.MaxLength = 255
        Me.cboInterBancaire.Name = "cboInterBancaire"
        Me.cboInterBancaire.Size = New System.Drawing.Size(142, 24)
        Me.cboInterBancaire.TabIndex = 4
        '
        'lblCommissionTransfert
        '
        Me.lblCommissionTransfert.Location = New System.Drawing.Point(18, 175)
        Me.lblCommissionTransfert.Name = "lblCommissionTransfert"
        Me.lblCommissionTransfert.Size = New System.Drawing.Size(256, 30)
        Me.lblCommissionTransfert.TabIndex = 5
        Me.lblCommissionTransfert.Text = "COMPTE  Commission sur Transfert_Ecobank"
        Me.lblCommissionTransfert.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'cboCommissionTransfert
        '
        Me.cboCommissionTransfert.FormattingEnabled = True
        Me.cboCommissionTransfert.Location = New System.Drawing.Point(280, 171)
        Me.cboCommissionTransfert.MaxLength = 255
        Me.cboCommissionTransfert.Name = "cboCommissionTransfert"
        Me.cboCommissionTransfert.Size = New System.Drawing.Size(142, 24)
        Me.cboCommissionTransfert.TabIndex = 6
        '
        'lblCommissionEnvoi
        '
        Me.lblCommissionEnvoi.Location = New System.Drawing.Point(18, 213)
        Me.lblCommissionEnvoi.Name = "lblCommissionEnvoi"
        Me.lblCommissionEnvoi.Size = New System.Drawing.Size(256, 30)
        Me.lblCommissionEnvoi.TabIndex = 7
        Me.lblCommissionEnvoi.Text = "COMPTE  Commission sur Envoi_Ecobank"
        Me.lblCommissionEnvoi.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'cboCommissionEnvoi
        '
        Me.cboCommissionEnvoi.FormattingEnabled = True
        Me.cboCommissionEnvoi.Location = New System.Drawing.Point(280, 209)
        Me.cboCommissionEnvoi.MaxLength = 255
        Me.cboCommissionEnvoi.Name = "cboCommissionEnvoi"
        Me.cboCommissionEnvoi.Size = New System.Drawing.Size(142, 24)
        Me.cboCommissionEnvoi.TabIndex = 8
        '
        'lblCommissionPaiement
        '
        Me.lblCommissionPaiement.Location = New System.Drawing.Point(18, 251)
        Me.lblCommissionPaiement.Name = "lblCommissionPaiement"
        Me.lblCommissionPaiement.Size = New System.Drawing.Size(256, 30)
        Me.lblCommissionPaiement.TabIndex = 9
        Me.lblCommissionPaiement.Text = "COMPTE  Commission sur Paiement_Ecobank"
        Me.lblCommissionPaiement.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'cboCommissionPaiement
        '
        Me.cboCommissionPaiement.FormattingEnabled = True
        Me.cboCommissionPaiement.Location = New System.Drawing.Point(280, 247)
        Me.cboCommissionPaiement.MaxLength = 255
        Me.cboCommissionPaiement.Name = "cboCommissionPaiement"
        Me.cboCommissionPaiement.Size = New System.Drawing.Size(142, 24)
        Me.cboCommissionPaiement.TabIndex = 10
        '
        'lblImpotsTaxeEnvoi
        '
        Me.lblImpotsTaxeEnvoi.Location = New System.Drawing.Point(18, 303)
        Me.lblImpotsTaxeEnvoi.Name = "lblImpotsTaxeEnvoi"
        Me.lblImpotsTaxeEnvoi.Size = New System.Drawing.Size(256, 30)
        Me.lblImpotsTaxeEnvoi.TabIndex = 11
        Me.lblImpotsTaxeEnvoi.Text = "COMPTE  IMPOTS ET TAXE SUR ENVOI"
        Me.lblImpotsTaxeEnvoi.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'cboImpotsTaxeEnvoi
        '
        Me.cboImpotsTaxeEnvoi.FormattingEnabled = True
        Me.cboImpotsTaxeEnvoi.Location = New System.Drawing.Point(280, 299)
        Me.cboImpotsTaxeEnvoi.MaxLength = 255
        Me.cboImpotsTaxeEnvoi.Name = "cboImpotsTaxeEnvoi"
        Me.cboImpotsTaxeEnvoi.Size = New System.Drawing.Size(142, 24)
        Me.cboImpotsTaxeEnvoi.TabIndex = 12
        '
        'lblTVA
        '
        Me.lblTVA.Location = New System.Drawing.Point(18, 341)
        Me.lblTVA.Name = "lblTVA"
        Me.lblTVA.Size = New System.Drawing.Size(256, 30)
        Me.lblTVA.TabIndex = 13
        Me.lblTVA.Text = "COMPTE TVA"
        Me.lblTVA.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'cboTVA
        '
        Me.cboTVA.FormattingEnabled = True
        Me.cboTVA.Location = New System.Drawing.Point(280, 337)
        Me.cboTVA.MaxLength = 255
        Me.cboTVA.Name = "cboTVA"
        Me.cboTVA.Size = New System.Drawing.Size(142, 24)
        Me.cboTVA.TabIndex = 14
        '
        'lblTTAEnvoi
        '
        Me.lblTTAEnvoi.Location = New System.Drawing.Point(18, 379)
        Me.lblTTAEnvoi.Name = "lblTTAEnvoi"
        Me.lblTTAEnvoi.Size = New System.Drawing.Size(256, 30)
        Me.lblTTAEnvoi.TabIndex = 15
        Me.lblTTAEnvoi.Text = "COMPTE  TTA sur Envoi WU"
        Me.lblTTAEnvoi.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'cboTTAEnvoi
        '
        Me.cboTTAEnvoi.FormattingEnabled = True
        Me.cboTTAEnvoi.Location = New System.Drawing.Point(280, 375)
        Me.cboTTAEnvoi.MaxLength = 255
        Me.cboTTAEnvoi.Name = "cboTTAEnvoi"
        Me.cboTTAEnvoi.Size = New System.Drawing.Size(142, 24)
        Me.cboTTAEnvoi.TabIndex = 16
        '
        'lblTTAReception
        '
        Me.lblTTAReception.Location = New System.Drawing.Point(18, 417)
        Me.lblTTAReception.Name = "lblTTAReception"
        Me.lblTTAReception.Size = New System.Drawing.Size(256, 30)
        Me.lblTTAReception.TabIndex = 17
        Me.lblTTAReception.Text = "COMPTE TTA sur paiement  WU"
        Me.lblTTAReception.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'cboTTAReception
        '
        Me.cboTTAReception.FormattingEnabled = True
        Me.cboTTAReception.Location = New System.Drawing.Point(280, 413)
        Me.cboTTAReception.MaxLength = 255
        Me.cboTTAReception.Name = "cboTTAReception"
        Me.cboTTAReception.Size = New System.Drawing.Size(142, 24)
        Me.cboTTAReception.TabIndex = 18
        '
        'btnEnregistrer
        '
        Me.btnEnregistrer.Location = New System.Drawing.Point(18, 465)
        Me.btnEnregistrer.Name = "btnEnregistrer"
        Me.btnEnregistrer.Size = New System.Drawing.Size(296, 38)
        Me.btnEnregistrer.TabIndex = 19
        Me.btnEnregistrer.Text = "Enregistrer"
        Me.btnEnregistrer.UseVisualStyleBackColor = True
        '
        'btnFermer
        '
        Me.btnFermer.Location = New System.Drawing.Point(322, 465)
        Me.btnFermer.Name = "btnFermer"
        Me.btnFermer.Size = New System.Drawing.Size(100, 38)
        Me.btnFermer.TabIndex = 20
        Me.btnFermer.Text = "Fermer"
        Me.btnFermer.UseVisualStyleBackColor = True
        '
        'lblInfo
        '
        Me.lblInfo.ForeColor = System.Drawing.SystemColors.GrayText
        Me.lblInfo.Location = New System.Drawing.Point(18, 511)
        Me.lblInfo.Name = "lblInfo"
        Me.lblInfo.Size = New System.Drawing.Size(404, 34)
        Me.lblInfo.TabIndex = 21
        Me.lblInfo.Text = ""
        '
        'FrmComptesSysteme
        '
        Me.AcceptButton = Me.btnEnregistrer
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnFermer
        Me.ClientSize = New System.Drawing.Size(434, 557)
        Me.Controls.Add(Me.panelTitre)
        Me.Controls.Add(Me.lblCompteCourant)
        Me.Controls.Add(Me.cboCompteCourant)
        Me.Controls.Add(Me.lblInterBancaire)
        Me.Controls.Add(Me.cboInterBancaire)
        Me.Controls.Add(Me.lblCommissionTransfert)
        Me.Controls.Add(Me.cboCommissionTransfert)
        Me.Controls.Add(Me.lblCommissionEnvoi)
        Me.Controls.Add(Me.cboCommissionEnvoi)
        Me.Controls.Add(Me.lblCommissionPaiement)
        Me.Controls.Add(Me.cboCommissionPaiement)
        Me.Controls.Add(Me.lblImpotsTaxeEnvoi)
        Me.Controls.Add(Me.cboImpotsTaxeEnvoi)
        Me.Controls.Add(Me.lblTVA)
        Me.Controls.Add(Me.cboTVA)
        Me.Controls.Add(Me.lblTTAEnvoi)
        Me.Controls.Add(Me.cboTTAEnvoi)
        Me.Controls.Add(Me.lblTTAReception)
        Me.Controls.Add(Me.cboTTAReception)
        Me.Controls.Add(Me.btnEnregistrer)
        Me.Controls.Add(Me.btnFermer)
        Me.Controls.Add(Me.lblInfo)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FrmComptesSysteme"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "COMPTE  Paramètres Systèmes"
        Me.panelTitre.ResumeLayout(False)
        Me.ResumeLayout(False)
    End Sub

    Friend WithEvents panelTitre As System.Windows.Forms.Panel
    Friend WithEvents lblTitre As System.Windows.Forms.Label
    Friend WithEvents lblCompteCourant As System.Windows.Forms.Label
    Friend WithEvents cboCompteCourant As System.Windows.Forms.ComboBox
    Friend WithEvents lblInterBancaire As System.Windows.Forms.Label
    Friend WithEvents cboInterBancaire As System.Windows.Forms.ComboBox
    Friend WithEvents lblCommissionTransfert As System.Windows.Forms.Label
    Friend WithEvents cboCommissionTransfert As System.Windows.Forms.ComboBox
    Friend WithEvents lblCommissionEnvoi As System.Windows.Forms.Label
    Friend WithEvents cboCommissionEnvoi As System.Windows.Forms.ComboBox
    Friend WithEvents lblCommissionPaiement As System.Windows.Forms.Label
    Friend WithEvents cboCommissionPaiement As System.Windows.Forms.ComboBox
    Friend WithEvents lblImpotsTaxeEnvoi As System.Windows.Forms.Label
    Friend WithEvents cboImpotsTaxeEnvoi As System.Windows.Forms.ComboBox
    Friend WithEvents lblTVA As System.Windows.Forms.Label
    Friend WithEvents cboTVA As System.Windows.Forms.ComboBox
    Friend WithEvents lblTTAEnvoi As System.Windows.Forms.Label
    Friend WithEvents cboTTAEnvoi As System.Windows.Forms.ComboBox
    Friend WithEvents lblTTAReception As System.Windows.Forms.Label
    Friend WithEvents cboTTAReception As System.Windows.Forms.ComboBox
    Friend WithEvents btnEnregistrer As System.Windows.Forms.Button
    Friend WithEvents btnFermer As System.Windows.Forms.Button
    Friend WithEvents lblInfo As System.Windows.Forms.Label

End Class
