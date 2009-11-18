Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.Xml

Public Class DefinitionWriter

    Private _Report As ReportDocument

#Region "Constructors"

    ''' <summary>
    ''' Use a report specified by file location
    ''' </summary>
    ''' <param name="filename"></param>
    ''' <remarks></remarks>
    Public Sub New(ByVal filename As String)

        _Report = New ReportDocument
        _Report.Load(filename, OpenReportMethod.OpenReportByTempCopy)

    End Sub

    ''' <summary>
    ''' Use a report reference
    ''' </summary>
    ''' <param name="value"></param>
    ''' <remarks></remarks>
    Public Sub New(ByRef value As ReportDocument)

        _Report = value

    End Sub

#End Region

#Region "Public Methods"

    Public Sub Create(ByVal destination As String)

        Using writer As XmlWriter = XmlWriter.Create(destination)

            writer.WriteStartDocument()
            writer.WriteStartElement("Report")

            ProcessReport(writer)

            writer.WriteEndElement()    '/Report
            writer.WriteEndDocument()

        End Using

    End Sub

#End Region

    Private Sub ProcessReport(ByRef writer As XmlWriter)

        With _Report
            writer.WriteAttributeString("FileName", .FileName.Replace("rassdk://", ""))
            writer.WriteAttributeString("HasSavedData", .HasSavedData)
            writer.WriteAttributeString("Name", .Name)
            'documented in the DataDefinition logic
            'writer.WriteElementString("RecordSelectionFormula", .RecordSelectionFormula)
        End With

        GetSummaryInfo(writer)
        GetReportOptions(writer)
        GetPrintOptions(writer)

        GetDatabase(writer)
        GetDataDefinition(writer)
        GetReportDefinition(writer)
        GetSubreports(writer)

    End Sub

#Region "General"

    Private Sub GetSummaryInfo(ByRef writer As XmlWriter)

        writer.WriteStartElement("SummaryInfo")
        With _Report.SummaryInfo

            writer.WriteAttributeString("KeywordsInReport", .KeywordsInReport)
            writer.WriteAttributeString("ReportAuthor", .ReportAuthor)
            writer.WriteAttributeString("ReportComments", .ReportComments)
            writer.WriteAttributeString("ReportSubject", .ReportSubject)
            writer.WriteAttributeString("ReportTitle", .ReportTitle)

        End With
        writer.WriteEndElement()    '/SummaryInfo

    End Sub

    Private Sub GetPrintOptions(ByRef writer As XmlWriter)

        writer.WriteStartElement("PrintOptions")
        With _Report.PrintOptions

            'writer.WriteAttributeString("CustomPaperSource", .CustomPaperSource.ToString)
            writer.WriteAttributeString("PageContentHeight", .PageContentHeight)
            writer.WriteAttributeString("PageContentWidth", .PageContentWidth)
            writer.WriteAttributeString("PaperOrientation", .PaperOrientation)
            writer.WriteAttributeString("PaperSize", .PaperSize)
            writer.WriteAttributeString("PaperSource", .PaperSource)
            writer.WriteAttributeString("PrinterDuplex", .PrinterDuplex)
            writer.WriteAttributeString("PrinterName", .PrinterName)

            writer.WriteStartElement("PageMargins")

            writer.WriteAttributeString("bottomMargin", .PageMargins.bottomMargin)
            writer.WriteAttributeString("leftMargin", .PageMargins.leftMargin)
            writer.WriteAttributeString("rightMargin", .PageMargins.rightMargin)
            writer.WriteAttributeString("topMargin", .PageMargins.topMargin)

            writer.WriteEndElement()    '/PageMargins

        End With
        writer.WriteEndElement()    '/PrintOptions

    End Sub

    Private Sub GetReportOptions(ByRef writer As XmlWriter)

        writer.WriteStartElement("ReportOptions")
        With _Report.ReportOptions
            writer.WriteAttributeString("EnableSaveDataWithReport", .EnableSaveDataWithReport)
            writer.WriteAttributeString("EnableSavePreviewPicture", .EnableSavePreviewPicture)
            writer.WriteAttributeString("EnableSaveSummariesWithReport", .EnableSaveSummariesWithReport)
            writer.WriteAttributeString("EnableUseDummyData", .EnableUseDummyData)
            writer.WriteAttributeString("InitialDataContext", .InitialDataContext)
            writer.WriteAttributeString("InitialReportPartName", .InitialReportPartName)
        End With
        writer.WriteEndElement()    '/ReportOptions

    End Sub

    Private Sub GetSubreports(ByRef writer As XmlWriter)

        writer.WriteStartElement("Reports")
        For Each subreport As ReportDocument In _Report.Subreports

            'recursively process each ReportDocument
            'ProcessReport(subreport, writer)

            'writer.WriteStartElement("Report")
            'With subreport
            '    writer.WriteAttributeString("EnableOnDemand", .EnableOnDemand)
            '    writer.WriteAttributeString("Height", .Height)
            '    writer.WriteAttributeString("Kind", .Kind.ToString)
            '    writer.WriteAttributeString("Left", .Left)
            '    writer.WriteAttributeString("Name", .Name)
            '    writer.WriteAttributeString("SubreportName", .SubreportName)
            '    writer.WriteAttributeString("Top", .Top)
            '    writer.WriteAttributeString("Width", .Width)
            'End With
            'writer.WriteEndElement()

        Next
        writer.WriteEndElement()    '/Subreports

    End Sub

#End Region

#Region "Database"

    Private Sub GetDatabase(ByRef writer As XmlWriter)

        writer.WriteStartElement("Database")

        GetTableLinks(writer)
        GetTables(writer)

        writer.WriteEndElement()    '/Database

    End Sub

    Private Sub GetTableLinks(ByRef writer As XmlWriter)

        writer.WriteStartElement("TableLinks")

        For Each TL As TableLink In _Report.Database.Links

            writer.WriteStartElement("TableLink")
            With TL

                writer.WriteAttributeString("JoinType", .JoinType)

                'source fields
                writer.WriteStartElement("SourceFields")

                For Each fd As FieldDefinition In TL.SourceFields
                    GetFieldDefinition(fd, writer)
                Next

                writer.WriteEndElement()    '/SourceFields

                'source table
                'GetTable(.SourceTable, writer)

                'destination fields
                writer.WriteStartElement("DestinationFields")

                For Each fd As FieldDefinition In TL.DestinationFields
                    GetFieldDefinition(fd, writer)
                Next

                writer.WriteEndElement()    '/DestinationFields

                'destination table
                'GetTable(.DestinationTable, writer)

            End With
            writer.WriteEndElement()    '/TableLink

        Next

        writer.WriteEndElement()    '/TableLinks

    End Sub

    Private Sub GetTables(ByRef writer As XmlWriter)

        writer.WriteStartElement("Tables")

        For Each T As Table In _Report.Database.Tables
            GetTable(T, writer)
        Next

        writer.WriteEndElement()    '/Tables

    End Sub

    Private Sub GetTable(ByRef table As Table, ByRef writer As XmlWriter)

        writer.WriteStartElement("Table")
        With table

            writer.WriteAttributeString("Location", .Location)
            writer.WriteAttributeString("Name", .Name)

            GetLogonInfo(.LogOnInfo, writer)

            writer.WriteStartElement("Fields")

            For Each fd As FieldDefinition In table.Fields
                GetFieldDefinition(fd, writer)
            Next

            writer.WriteEndElement()    '/Fields

        End With
        writer.WriteEndElement()    '/Table

    End Sub

    Private Sub GetLogonInfo(ByRef li As TableLogOnInfo, ByRef writer As XmlWriter)

        writer.WriteStartElement("LogonInfo")
        With li

            writer.WriteAttributeString("ReportName", .ReportName)
            writer.WriteAttributeString("TableName", .TableName)

            GetConnectionInfo(.ConnectionInfo, writer)

        End With
        writer.WriteEndElement()    '/LogonInfo

    End Sub

    'todo: add logic to handle interactive database prompting if necessary
    Private Sub GetConnectionInfo(ByVal ci As ConnectionInfo, ByRef writer As XmlWriter)

        writer.WriteStartElement("ConnectionInfo")
        With ci

            writer.WriteAttributeString("AllowCustomConnection", .AllowCustomConnection)
            'writer.WriteAttributeString("EncodedPassword", .EncodedPassword)
            writer.WriteAttributeString("DatabaseName", .DatabaseName)
            writer.WriteAttributeString("IntegratedSecurity", .IntegratedSecurity)
            writer.WriteAttributeString("Password", .Password)
            writer.WriteAttributeString("ServerName", .ServerName)
            writer.WriteAttributeString("Type", .Type.ToString)
            writer.WriteAttributeString("UserID", .UserID)

        End With
        writer.WriteEndElement()    '/ConnectionInfo

    End Sub

    Private Sub GetFieldDefinition(ByRef fd As FieldDefinition, ByRef writer As XmlWriter)

        writer.WriteStartElement("Field")
        With fd

            writer.WriteAttributeString("FormulaName", .FormulaName)
            writer.WriteAttributeString("Kind", .Kind.ToString)
            writer.WriteAttributeString("Name", .Name)
            writer.WriteAttributeString("NumberOfBytes", .NumberOfBytes)
            writer.WriteAttributeString("UseCount", .UseCount)
            writer.WriteAttributeString("ValueType", .ValueType)

        End With
        writer.WriteEndElement()    '/Field

    End Sub

#End Region

#Region "Data Definition"

    Private Sub GetDataDefinition(ByRef writer As XmlWriter)

        writer.WriteStartElement("DataDefinition")
        With _Report.DataDefinition

            writer.WriteStartElement("FormulaFieldDefinitions")
            For Each field In .FormulaFields
                GetFieldObject(field, writer)
            Next
            writer.WriteEndElement()

            writer.WriteStartElement("GroupNameFieldDefinitions")
            For Each field In .GroupNameFields
                GetFieldObject(field, writer)
            Next
            writer.WriteEndElement()

            writer.WriteStartElement("Groups")
            For Each field In .Groups
                '    GetFieldObject(field, writer)
            Next
            writer.WriteEndElement()

            writer.WriteElementString("GroupSelectionFormula", .GroupSelectionFormula)

            writer.WriteStartElement("ParameterFieldDefinitions")
            For Each field In .ParameterFields
                GetFieldObject(field, writer)
            Next
            writer.WriteEndElement()

            writer.WriteElementString("RecordSelectionFormula", .RecordSelectionFormula)

            writer.WriteStartElement("RunningTotalFieldDefinitions")
            For Each field In .RunningTotalFields
                GetFieldObject(field, writer)
            Next
            writer.WriteEndElement()

            writer.WriteStartElement("SortFields")
            'For Each field In .SortFields
            '    GetFieldObject(field, writer)
            'Next
            writer.WriteEndElement()

            writer.WriteStartElement("SQLExpressionFields")
            For Each field In .SQLExpressionFields
                GetFieldObject(field, writer)
            Next
            writer.WriteEndElement()

            writer.WriteStartElement("SummaryFields")
            For Each field In .SummaryFields
                GetFieldObject(field, writer)
            Next
            writer.WriteEndElement()

        End With
        writer.WriteEndElement()    '/DataDefinition

    End Sub

#End Region

#Region "Report Definition"

    Private Sub GetReportDefinition(ByRef writer As XmlWriter)

        writer.WriteStartElement("ReportDefinition")
        writer.WriteStartElement("Areas")

        For Each A As Area In _Report.ReportDefinition.Areas
            GetArea(A, writer)
        Next

        writer.WriteEndElement()    '/Areas
        writer.WriteEndElement()    '/ReportDefinition

    End Sub

    Private Sub GetArea(ByRef a As Area, ByRef writer As XmlWriter)

        writer.WriteStartElement("Area")
        With a

            writer.WriteAttributeString("Kind", .Kind.ToString)
            writer.WriteAttributeString("Name", .Name)

            writer.WriteStartElement("AreaFormat")
            With .AreaFormat

                writer.WriteAttributeString("EnableHideForDrillDown", .EnableHideForDrillDown)
                writer.WriteAttributeString("EnableKeepTogether", .EnableKeepTogether)
                writer.WriteAttributeString("EnableNewPageAfter", .EnableNewPageAfter)
                writer.WriteAttributeString("EnableNewPageBefore", .EnableNewPageBefore)
                writer.WriteAttributeString("EnablePrintAtBottomOfPage", .EnablePrintAtBottomOfPage)
                writer.WriteAttributeString("EnableResetPageNumberAfter", .EnableResetPageNumberAfter)
                writer.WriteAttributeString("EnableSuppress", .EnableSuppress)

            End With
            writer.WriteEndElement()    '/AreaFormat

            writer.WriteStartElement("Sections")

            For Each s As Section In a.Sections
                GetSection(s, writer)
            Next

            writer.WriteEndElement()    '/Sections

        End With
        writer.WriteEndElement()    '/Area

    End Sub

    Private Sub GetSection(ByRef section As Section, ByRef writer As XmlWriter)

        writer.WriteStartElement("Section")
        With section

            writer.WriteAttributeString("Height", .Height)
            writer.WriteAttributeString("Kind", .Kind.ToString)
            writer.WriteAttributeString("Name", .Name)

            writer.WriteStartElement("SectionFormat")
            With .SectionFormat

                'writer.WriteAttributeString("BackgroundColor", .BackgroundColor)
                writer.WriteAttributeString("CssClass", .CssClass)
                writer.WriteAttributeString("EnableKeepTogether", .EnableKeepTogether)
                writer.WriteAttributeString("EnableNewPageAfter", .EnableNewPageAfter)
                writer.WriteAttributeString("EnableNewPageBefore", .EnableNewPageBefore)
                writer.WriteAttributeString("EnablePrintAtBottomOfPage", .EnablePrintAtBottomOfPage)
                writer.WriteAttributeString("EnableResetPageNumberAfter", .EnableResetPageNumberAfter)
                writer.WriteAttributeString("EnableSuppress", .EnableSuppress)
                writer.WriteAttributeString("EnableSuppressIfBlank", .EnableSuppressIfBlank)
                writer.WriteAttributeString("EnableUnderlaySection", .EnableUnderlaySection)

            End With
            writer.WriteEndElement()    '/SectionFormat

            writer.WriteStartElement("ReportObjects")

            For Each ro As ReportObject In .ReportObjects
                GetReportObject(ro, writer)
            Next

            writer.WriteEndElement()    '/ReportObjects

        End With
        writer.WriteEndElement()    '/Section

    End Sub

    'todo: capture formatting details (e.g. font, conditional formulae) of each kind of object
    Private Sub GetReportObject(ByRef reportObject As Object, ByRef writer As XmlWriter)

        If TypeOf (reportobject) Is BlobFieldObject Then
            writer.WriteStartElement("BlobFieldObject")
            writer.WriteEndElement()

        ElseIf TypeOf (reportobject) Is ChartObject Then
            writer.WriteStartElement("ChartObject")
            writer.WriteEndElement()

        ElseIf TypeOf (reportobject) Is CrossTabObject Then
            writer.WriteStartElement("CrossTabObject")
            writer.WriteEndElement()

        ElseIf TypeOf (reportobject) Is DrawingObject Then
            writer.WriteStartElement("DrawingObject")
            writer.WriteEndElement()

            'todo: generate enough information about FieldObject to be able to reference object in DataDefinition section
        ElseIf TypeOf (reportobject) Is FieldObject Then
            writer.WriteStartElement("FieldObject")
            writer.WriteEndElement()

        ElseIf TypeOf (reportobject) Is GraphicObject Then
            writer.WriteStartElement("GraphicObject")
            writer.WriteEndElement()

        ElseIf TypeOf (reportobject) Is MapObject Then
            writer.WriteStartElement("MapObject")
            writer.WriteEndElement()

        ElseIf TypeOf (reportobject) Is OlapGridObject Then
            writer.WriteStartElement("OlapGridObject")
            writer.WriteEndElement()

            'todo: recursively process subreport
        ElseIf TypeOf (reportobject) Is SubreportObject Then
            Dim subreportObject As SubreportObject = CType(reportobject, SubreportObject)
            'Dim subreportName As String = sr.SubreportName
            '_Report.OpenSubreport(subreportName)

            writer.WriteStartElement("SubreportObject")
            With subreportObject
                writer.WriteAttributeString("EnableOnDemand", .EnableOnDemand)
                writer.WriteAttributeString("Height", .Height)
                writer.WriteAttributeString("Kind", .Kind.ToString)
                writer.WriteAttributeString("Left", .Left)
                writer.WriteAttributeString("Name", .Name)
                writer.WriteAttributeString("SubreportName", .SubreportName)
                writer.WriteAttributeString("Top", .Top)
                writer.WriteAttributeString("Width", .Width)
            End With
            writer.WriteEndElement()

        ElseIf TypeOf (reportobject) Is TextObject Then
            writer.WriteStartElement("TextObject")
            writer.WriteEndElement()

        End If

    End Sub

    'todo: capture additional properties of each FieldDefinition
    Private Sub GetFieldObject(ByRef fo As Object, ByRef writer As XmlWriter)

        If TypeOf fo Is DatabaseFieldDefinition Then

            Dim DF As DatabaseFieldDefinition = DirectCast(fo, DatabaseFieldDefinition)

            writer.WriteStartElement("DatabaseFieldDefinition")
            With DF
                writer.WriteAttributeString("FormulaName", .FormulaName)
                writer.WriteAttributeString("Kind", .Kind.ToString)
                writer.WriteAttributeString("Name", .Name)
                writer.WriteAttributeString("NumberOfBytes", .NumberOfBytes)
                writer.WriteAttributeString("TableName", .TableName)
                writer.WriteAttributeString("UseCount", .UseCount)
                writer.WriteAttributeString("ValueType", .ValueType.ToString)
            End With
            writer.WriteEndElement()

        ElseIf TypeOf fo Is FormulaFieldDefinition Then

            Dim FF As FormulaFieldDefinition = DirectCast(fo, FormulaFieldDefinition)

            writer.WriteStartElement("FormulaFieldDefinition")
            With FF
                writer.WriteAttributeString("FormulaName", .FormulaName)
                writer.WriteAttributeString("Kind", .Kind.ToString)
                writer.WriteAttributeString("Name", .Name)
                writer.WriteAttributeString("NumberOfBytes", .NumberOfBytes)
                writer.WriteAttributeString("Text", .Text)
                writer.WriteAttributeString("UseCount", .UseCount)
                writer.WriteAttributeString("ValueType", .ValueType.ToString)
            End With
            writer.WriteEndElement()

        ElseIf TypeOf fo Is GroupNameFieldDefinition Then

            Dim GNF As GroupNameFieldDefinition = DirectCast(fo, GroupNameFieldDefinition)

            writer.WriteStartElement("GroupNameFieldDefinition")
            With GNF
                writer.WriteAttributeString("FormulaName", .FormulaName)
                writer.WriteAttributeString("Group", .Group.ToString)
                writer.WriteAttributeString("GroupNameFieldName", .GroupNameFieldName)
                writer.WriteAttributeString("Kind", .Kind.ToString)
                writer.WriteAttributeString("Name", .Name)
                writer.WriteAttributeString("NumberOfBytes", .NumberOfBytes)
                writer.WriteAttributeString("UseCount", .UseCount)
                writer.WriteAttributeString("ValueType", .ValueType.ToString)
            End With
            writer.WriteEndElement()

        ElseIf TypeOf fo Is ParameterFieldDefinition Then

            Dim PF As ParameterFieldDefinition = DirectCast(fo, ParameterFieldDefinition)

            writer.WriteStartElement("ParameterFieldDefinition")
            With PF
                '.CurrentValues
                'writer.WriteStartElement("CurrentValues")
                'writer.WriteEndElement()   '/CurrentValues
                '.DefaultValueDisplayType
                '.DefaultValues
                'writer.WriteStartElement("DefaultValues")
                'writer.WriteEndElement()   '/DefaultValues
                '.DefaultValueSortMethod
                '.DefaultValueSortOrder
                '.DiscreteOrRangeKind
                writer.WriteAttributeString("EditMask", .EditMask)
                writer.WriteAttributeString("EnableAllowEditingDefaultValue", .EnableAllowEditingDefaultValue)
                writer.WriteAttributeString("EnableAllowMultipleValue", .EnableAllowMultipleValue)
                writer.WriteAttributeString("EnableNullValue", .EnableNullValue)
                writer.WriteAttributeString("FormulaName", .FormulaName)
                writer.WriteAttributeString("HasCurrentValue", .HasCurrentValue)
                writer.WriteAttributeString("Kind", .Kind.ToString)
                writer.WriteAttributeString("MaximumValue", .MaximumValue)
                writer.WriteAttributeString("MinimumValue", .MinimumValue)
                writer.WriteAttributeString("Name", .Name)
                writer.WriteAttributeString("NumberOfBytes", .NumberOfBytes)
                writer.WriteAttributeString("ParameterFieldName", .ParameterFieldName)
                writer.WriteAttributeString("ParameterFieldUsage", .ParameterFieldUsage.ToString)
                writer.WriteAttributeString("ParameterType", .ParameterType.ToString)
                writer.WriteAttributeString("ParameterValueKind", .ParameterValueKind.ToString)
                writer.WriteAttributeString("PromptText", .PromptText)
                writer.WriteAttributeString("ReportName", .ReportName)
                writer.WriteAttributeString("UseCount", .UseCount)
                writer.WriteAttributeString("ValueType", .ValueType.ToString)
            End With
            writer.WriteEndElement()

        ElseIf TypeOf fo Is RunningTotalFieldDefinition Then

            Dim RTF As RunningTotalFieldDefinition = DirectCast(fo, RunningTotalFieldDefinition)

            writer.WriteStartElement("RunningTotalFieldDefinition")
            With RTF
                writer.WriteAttributeString("EvaluationConditionType", .EvaluationCondition)
                writer.WriteAttributeString("EvaluationConditionType", .EvaluationConditionType.ToString)
                writer.WriteAttributeString("FormulaName", .FormulaName)
                writer.WriteAttributeString("Group", .Group.ToString)
                writer.WriteAttributeString("Kind", .Kind.ToString)
                writer.WriteAttributeString("Name", .Name)
                writer.WriteAttributeString("NumberOfBytes", .NumberOfBytes)
                writer.WriteAttributeString("Operation", .Operation)
                writer.WriteAttributeString("OperationParameter", .OperationParameter)
                writer.WriteAttributeString("ResetCondition", .ResetCondition)
                writer.WriteAttributeString("ResetConditionType", .ResetConditionType.ToString)
                writer.WriteAttributeString("SecondarySummarizedField", .SecondarySummarizedField.FormulaName)
                writer.WriteAttributeString("SummarizedField", .SummarizedField.FormulaName)
                writer.WriteAttributeString("UseCount", .UseCount)
                writer.WriteAttributeString("ValueType", .ValueType.ToString)
            End With
            writer.WriteEndElement()

        ElseIf TypeOf fo Is SpecialVarFieldDefinition Then

            writer.WriteStartElement("SpecialVarFieldDefinition")
            Dim SVF As SpecialVarFieldDefinition = DirectCast(fo, SpecialVarFieldDefinition)
            With SVF
                writer.WriteAttributeString("FormulaName", .FormulaName)
                writer.WriteAttributeString("Kind", .Kind.ToString)
                writer.WriteAttributeString("Name", .Name)
                writer.WriteAttributeString("NumberOfBytes", .NumberOfBytes)
                writer.WriteAttributeString("SpecialVarType", .SpecialVarType.ToString())
                writer.WriteAttributeString("UseCount", .UseCount)
                writer.WriteAttributeString("ValueType", .ValueType.ToString)
            End With
            writer.WriteEndElement()

        ElseIf TypeOf fo Is SQLExpressionFieldDefinition Then

            writer.WriteStartElement("SQLExpressionFieldDefinition")
            Dim SEF As SQLExpressionFieldDefinition = DirectCast(fo, SQLExpressionFieldDefinition)
            With SEF
                writer.WriteAttributeString("FormulaName", .FormulaName)
                writer.WriteAttributeString("Kind", .Kind.ToString)
                writer.WriteAttributeString("Name", .Name)
                writer.WriteAttributeString("NumberOfBytes", .NumberOfBytes)
                writer.WriteAttributeString("Text", .Text)
                writer.WriteAttributeString("UseCount", .UseCount)
                writer.WriteAttributeString("ValueType", .ValueType.ToString)
            End With
            writer.WriteEndElement()

        ElseIf TypeOf fo Is SummaryFieldDefinition Then

            writer.WriteStartElement("SummaryFieldDefinition")
            Dim SF As SummaryFieldDefinition = DirectCast(fo, SummaryFieldDefinition)
            With SF
                writer.WriteAttributeString("FormulaName", .FormulaName)
                writer.WriteAttributeString("Group", .Group.ToString)
                writer.WriteAttributeString("Kind", .Kind.ToString)
                writer.WriteAttributeString("Name", .Name)
                writer.WriteAttributeString("NumberOfBytes", .NumberOfBytes)
                writer.WriteAttributeString("Operation", .Operation.ToString)
                writer.WriteAttributeString("OperationParameter", .OperationParameter.ToString)
                writer.WriteAttributeString("SecondarySummarizedField", .SecondarySummarizedField.ToString)
                writer.WriteAttributeString("SummarizedField", .SummarizedField.ToString)
                writer.WriteAttributeString("UseCount", .UseCount)
                writer.WriteAttributeString("ValueType", .ValueType.ToString)
            End With
            writer.WriteEndElement()

        End If

    End Sub

#End Region

End Class
