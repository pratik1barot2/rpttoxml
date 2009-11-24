Imports RptToXml

Module Module1

    Sub Main()

        Dim path As String = "C:\Documents and Settings\Administrator\My Documents\Visual Studio 2008\Projects\RptToXml\Harness\"
        Dim fileName As String = "Sample.rpt"
        'Dim fileName As String = "Employee Profile.rpt"
        'Dim fileName As String = "CD-0130-100010 Open Medical Record Audit.006.rpt"
        Dim writer As New RptDefinitionWriter(path & fileName)
        writer.ShowFormatTypes = RptDefinitionWriter.FormatTypes.None
        writer.WriteToXml(path & fileName.Replace(".rpt", ".xml"))

    End Sub

End Module
