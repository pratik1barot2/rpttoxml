Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.Xml

Public Class RptDefinitionReader

#Region "Public Methods"

    ''' <summary>
    ''' Saves the report-definition document to the specified file.
    ''' </summary>
    ''' <param name="targetRptPath">The location of the file where you want to save the document.</param>
    ''' <remarks></remarks>
    Public Sub WriteToRpt(ByVal targetRptPath As String)

        Throw New NotImplementedException

    End Sub

#End Region

#Region "Private Methods"

    Private Sub ProcessXml(ByRef report As ReportDocument, ByRef reader As XmlReader)

    End Sub

#End Region

End Class
