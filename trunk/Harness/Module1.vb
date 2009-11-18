Imports RptToXml

Module Module1

    Sub Main()

        Dim fileName As String = "C:\Documents and Settings\Administrator\My Documents\Visual Studio 2008\Projects\RptToXml\Harness\Sample.rpt"
        'Dim fileName As String = "C:\Documents and Settings\Administrator\My Documents\Visual Studio 2008\Projects\RptToXml\Harness\Employee Profile.rpt"
        Dim writer As New DefinitionWriter(fileName)
        writer.Create(fileName.Replace(".rpt", ".xml"))

    End Sub

End Module
