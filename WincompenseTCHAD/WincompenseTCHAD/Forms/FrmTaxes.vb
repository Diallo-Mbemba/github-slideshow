Option Strict On
Option Explicit On

Imports System.Windows.Forms

''' <summary>
''' Affiche le barème du produit en service : ses taxes, leurs taux, leurs assiettes et les
''' comptes qu'elles créditent.
'''
''' CET ÉCRAN NE MODIFIE RIEN, ET CE N'EST PAS UN OUBLI. Le barème de Western Union est
''' réconcilié avec la banque ; la banque a décidé qu'il ne se changerait pas depuis une
''' fenêtre. Voir BaremeTaxesWU pour le raisonnement complet. L'écran ouvrira la saisie le
''' jour où un produit modifiable sera en service — Ria d'abord — et c'est
''' BaremeTaxesWU.Modifiable qui le lui dira, à un seul endroit.
'''
''' CE QU'IL APPORTE DÈS AUJOURD'HUI. Le barème n'était lisible que dans le code : cinq taux,
''' cinq comptes, QUATRE assiettes différentes et un résidu partagé en deux. Personne à la
''' banque ne pouvait le vérifier sans ouvrir Visual Studio — au moment même où elle nous
''' contestait deux lignes de pièce. Le rendre lisible, c'est déjà répondre.
'''
''' AUCUNE DÉPENDANCE AU CALCUL. Ce formulaire lit BaremeTaxesWU et rien d'autre ; aucun
''' service de calcul ne le lit en retour. On peut le supprimer sans qu'une compensation
''' change d'un franc.
''' </summary>
Public Class FrmTaxes

    ''' <summary>Noms des colonnes de la grille. Écrits une fois, employés partout.</summary>
    Private Const COL_TAXE As String = "Taxe"
    Private Const COL_TAUX As String = "Taux"
    Private Const COL_ASSIETTE As String = "Assiette"
    Private Const COL_PIECE As String = "Ligne sur la pièce"
    Private Const COL_COMPTE As String = "Compte"

    Public Sub New()
        InitializeComponent()
        IconesWU.Habiller(Me)
    End Sub

#Region "Chargement"

    Private Sub FrmTaxes_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim bareme As BaremeTaxesWU = BaremeTaxesWU.Actuel

        lblTitre.Text = "Taxes et barème"
        Me.Text = $"Taxes et barème — {bareme.Produit}"

        lblIntro.Text =
            $"Barème appliqué à la compensation {bareme.Produit}. Chaque taux porte sur SON assiette : " &
            "les charges d'envoi, le principal, ou le solde de taxes — ce n'est pas la même pour toutes." &
            Environment.NewLine &
            "Les comptes viennent du paramétrage en service et se changent dans Paramétrage > Comptes systèmes."

        AfficherLeVerrou(bareme)
        RemplirLaGrille(bareme)

        lblFormule.Text = bareme.FormuleDuSolde()
        lblNotes.Text = RedigerLesNotes(bareme)
    End Sub

    ''' <summary>
    ''' Dit si le barème se modifie ici, et pourquoi.
    '''
    ''' Le bandeau est posé même quand le barème EST modifiable : un écran qui ne parle que
    ''' pour refuser apprend à ne plus être lu. Sa couleur change, son texte aussi ; sa place
    ''' ne bouge pas.
    ''' </summary>
    Private Sub AfficherLeVerrou(bareme As BaremeTaxesWU)

        If bareme.Modifiable Then

            panelVerrou.BackColor = Drawing.Color.FromArgb(232, 245, 233)
            lblVerrou.ForeColor = Drawing.Color.FromArgb(27, 94, 32)
            lblVerrou.Text =
                $"Barème {bareme.Produit} — modifiable." & Environment.NewLine &
                "Les taux sont conservés en base et valent pour tous les postes."
            Return
        End If

        panelVerrou.BackColor = Drawing.Color.FromArgb(255, 248, 225)
        lblVerrou.ForeColor = Drawing.Color.FromArgb(120, 80, 0)
        lblVerrou.Text =
            $"Barème {bareme.Produit} — figé." & Environment.NewLine &
            "Ces taux sont réconciliés avec la banque et ne se modifient pas ici." &
            Environment.NewLine &
            "Les produits ajoutés plus tard auront leur barème modifiable."
    End Sub

    ''' <summary>
    ''' Remplit la grille. Les colonnes sont construites ici, et non liées à un objet : la
    ''' grille montre le barème TEL QU'ON LE LIT — un taux en pourcentage, une assiette en
    ''' clair — et non les champs tels qu'ils sont rangés.
    ''' </summary>
    Private Sub RemplirLaGrille(bareme As BaremeTaxesWU)

        Dim table As New DataTable("Bareme")
        table.Columns.Add(COL_TAXE, GetType(String))
        table.Columns.Add(COL_TAUX, GetType(String))
        table.Columns.Add(COL_ASSIETTE, GetType(String))
        table.Columns.Add(COL_PIECE, GetType(String))
        table.Columns.Add(COL_COMPTE, GetType(String))

        For Each taxe As TaxeWU In bareme.Lignes

            Dim assiette As String = taxe.LibelleDeLAssiette
            If taxe.SousAgentSeulement Then assiette &= " — sous-agents seulement"

            table.Rows.Add(taxe.Libelle,
                           taxe.TauxEnTexte(),
                           assiette,
                           If(taxe.LignePosee, "oui", "non — absorbée par la contrepartie"),
                           If(String.IsNullOrWhiteSpace(taxe.Compte), "(non paramétré)", taxe.Compte))
        Next

        dgvTaxes.AutoGenerateColumns = True
        dgvTaxes.DataSource = table
        dgvTaxes.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

        ' Le numéro de compte se lit chiffre par chiffre : une police à chasse fixe évite de
        ' confondre deux comptes qui ne diffèrent que par un caractère.
        If dgvTaxes.Columns.Contains(COL_COMPTE) Then
            dgvTaxes.Columns(COL_COMPTE).DefaultCellStyle.Font = New Drawing.Font("Consolas", 9.0!)
            dgvTaxes.Columns(COL_COMPTE).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        End If

        GrilleWU.LargeurFixe(dgvTaxes, COL_TAUX, 80)
        GrilleWU.LargeurFixe(dgvTaxes, COL_COMPTE, 110)

        ' Une grille en lecture seule qui s'ouvre sur une ligne surlignée laisse croire qu'un
        ' choix a été fait. Il n'y a rien à choisir ici.
        dgvTaxes.ClearSelection()
    End Sub

    ''' <summary>
    ''' Les remarques que la grille ne peut pas porter : ce qui se calcule par différence, ce
    ''' qui n'est pas une taxe, et ce qui ne figure pas dans ce barème.
    '''
    ''' Elles sont rédigées À PARTIR DU BARÈME, et non recopiées : une note qui décrirait une
    ''' ligne disparue serait pire que pas de note du tout.
    ''' </summary>
    Private Shared Function RedigerLesNotes(bareme As BaremeTaxesWU) As String

        Dim notes As New System.Text.StringBuilder()

        notes.AppendLine("Le TOTAL DES TAXES n'est pas calculé : il est lu dans le rapport, c'est ce que " &
                         "le produit a déjà prélevé au client. Les deux quotes-parts se partagent ce qu'il en reste.")

        Dim produit As TaxeWU = bareme.Lignes.FirstOrDefault(Function(t) t.EstUnProduitBancaire)

        If produit IsNot Nothing Then
            notes.AppendLine($"« {produit.Libelle} » n'est pas une taxe : son compte {produit.Compte} est un " &
                             "compte de PRODUIT bancaire. Elle figure ici parce qu'elle se calcule sur le solde " &
                             "de taxes, comme sa jumelle, et qu'on ne comprend ni l'une ni l'autre séparément.")
        End If

        Dim reserve As TaxeWU = bareme.Lignes.FirstOrDefault(Function(t) t.SousAgentSeulement)

        If reserve IsNot Nothing Then
            notes.AppendLine($"« {reserve.Libelle} » n'est pas due par une agence propre — règle confirmée par " &
                             "la banque. Elle n'est pas calculée puis omise : elle n'est pas due.")
        End If

        notes.Append("La commission sur envoi, " &
                     (ConstantesWU.TAUX_COMMISSION_ENVOI * 100D).ToString("0.###", Globalization.CultureInfo.CurrentCulture) &
                     " % des charges, ne figure pas dans ce barème : c'est un produit de la banque, pas une taxe.")

        Return notes.ToString()
    End Function

#End Region

#Region "Commandes"

    Private Sub btnFermer_Click(sender As Object, e As EventArgs) Handles btnFermer.Click
        Close()
    End Sub

#End Region

End Class
