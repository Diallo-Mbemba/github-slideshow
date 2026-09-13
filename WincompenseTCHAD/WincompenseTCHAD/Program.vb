Option Strict On
Option Explicit On

Imports System.Windows.Forms

''' <summary>Point d'entrée de l'application Windows Forms.</summary>
Module Program

    <STAThread()>
    Sub Main()
        Application.EnableVisualStyles()
        Application.SetCompatibleTextRenderingDefault(False)
        ' La fenêtre MDI est le point d'entrée : elle ouvre elle-même l'écran de traitement.
        Application.Run(New FrmPrincipal())
    End Sub

End Module
