Option Strict On
Option Explicit On

Imports System.Windows.Forms

''' <summary>Point d'entrée de l'application Windows Forms.</summary>
Module Program

    <STAThread()>
    Sub Main()
        Application.EnableVisualStyles()
        Application.SetCompatibleTextRenderingDefault(False)
        Application.Run(New FrmCompensationWU())
    End Sub

End Module
