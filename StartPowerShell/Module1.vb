Module Module1

    ReadOnly v_Path As String = Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments) & "\ServiceService"

    Sub Main()
        Process.Start("powershell", String.Format("-noexit {0}\P.ps1", v_Path))
    End Sub

End Module
