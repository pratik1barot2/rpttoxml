Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.Xml

Public Class RptDefinitionWriter

    'http://dotnet.org.za/kevint/articles/Flags.aspx
    <Flags()> _
    Public Enum FormatTypes
        None = 0
        Border = 2 ^ 0
        Color = 2 ^ 1
        Font = 2 ^ 2
        AreaFormat = 2 ^ 3
        FieldFormat = 2 ^ 4
        ObjectFormat = 2 ^ 5
        SectionFormat = 2 ^ 6
        All = Border And Color And Font And AreaFormat And FieldFormat And ObjectFormat And SectionFormat
    End Enum

    '<Flags()> _
    'Public Enum ObjectTypes
    '    None = 0
    '    Area = 2 ^ 0
    '    Section = 2 ^ 1
    '    ReportObject = 2 ^ 2
    '    All = Area And Section And ReportObject
    'End Enum

    Private _Report As ReportDocument

    Private _ShowFormatTypes As FormatTypes = FormatTypes.AreaFormat Or FormatTypes.SectionFormat Or FormatTypes.Color
    'Private _ShowObjectTypes As ObjectTypes = ObjectTypes.ReportObject

#Region "Properties"

    Private Property Report() As ReportDocument
        Get
            Return _Report
        End Get
        Set(ByVal value As ReportDocument)
            _Report = value
        End Set
    End Property

    Public Property ShowFormatTypes() As FormatTypes
        Get
            Return _ShowFormatTypes
        End Get
        Set(ByVal value As FormatTypes)
            _ShowFormatTypes = value
        End Set
    End Property

    'Public Property ShowObjectTypes() As ObjectTypes
    '    Get
    '        Return _ShowObjectTypes
    '    End Get
    '    Set(ByVal value As ObjectTypes)
    '        _ShowObjectTypes = value
    '    End Set
    'End Property

#End Region

#Region "Constructors"

    ''' <summary>
    ''' Use a report specified by file location
    ''' </summary>
    ''' <param name="filename"></param>
    ''' <remarks></remarks>
    Public Sub New(ByVal filename As String)

        Me.Report = New ReportDocument

        Try
            Me.Report.Load(filename, OpenReportMethod.OpenReportByTempCopy)

        Catch ex As Exception
            'preserve the stack trace information
            Throw

        Finally
            'Me.Report.Close()

        End Try

    End Sub

    ''' <summary>
    ''' Use a report reference
    ''' </summary>
    ''' <param name="value"></param>
    ''' <remarks></remarks>
    Public Sub New(ByRef value As ReportDocument)

        Me.Report = value

    End Sub

#End Region

#Region "Public Methods"

    ''' <summary>
    ''' 	Saves the report-definition document to the specified stream.
    ''' </summary>
    ''' <param name="output">The stream to which you want to save.</param>
    ''' <remarks></remarks>
    Public Sub WriteToXml(ByVal output As System.IO.Stream)

        WriteToXml(XmlWriter.Create(output))

    End Sub

    ''' <summary>
    ''' Saves the report-definition document to the specified file.
    ''' </summary>
    ''' <param name="targetXmlPath">The location of the file where you want to save the document.</param>
    ''' <remarks></remarks>
    Public Sub WriteToXml(ByVal targetXmlPath As String)

        WriteToXml(XmlWriter.Create(targetXmlPath))

    End Sub

    ''' <summary>
    ''' Saves the report-definition document to the specified XmlWriter.
    ''' </summary>
    ''' <param name="writer">The XmlWriter to which you want to save.</param>
    ''' <remarks></remarks>
    Public Sub WriteToXml(ByRef writer As XmlWriter)

        Using writer

            writer.WriteStartDocument()

            ProcessReport(Me.Report, writer)

            writer.WriteEndDocument()

        End Using

    End Sub

#End Region

#Region "Private Methods"

    Private Sub ProcessReport(ByRef report As ReportDocument, ByRef writer As XmlWriter)

        writer.WriteStartElement("Report")
        With report

            writer.WriteAttributeString("Name", .Name)

            If Not .IsSubreport Then
                writer.WriteAttributeString("FileName", .FileName.Replace("rassdk://", ""))
                writer.WriteAttributeString("HasSavedData", .HasSavedData)

                GetSummaryInfo(report, writer)
                GetReportOptions(report, writer)
                GetPrintOptions(report, writer)
                GetSubreports(report, writer)

            End If

        End With

        GetDatabase(report, writer)
        GetDataDefinition(report, writer)
        GetReportDefinition(report, writer)

        writer.WriteEndElement()    '/Report

    End Sub

#End Region

#Region "Main Report Only"

    Private Sub GetSummaryInfo(ByRef report As ReportDocument, ByRef writer As XmlWriter)

        writer.WriteStartElement("SummaryInfo")
        With report.SummaryInfo

            writer.WriteAttributeString("KeywordsInReport", .KeywordsInReport)
            writer.WriteAttributeString("ReportAuthor", .ReportAuthor)
            writer.WriteAttributeString("ReportComments", .ReportComments)
            writer.WriteAttributeString("ReportSubject", .ReportSubject)
            writer.WriteAttributeString("ReportTitle", .ReportTitle)

        End With
        writer.WriteEndElement()    '/SummaryInfo

    End Sub

    Private Sub GetReportOptions(ByRef report As ReportDocument, ByRef writer As XmlWriter)

        writer.WriteStartElement("ReportOptions")
        With report.ReportOptions
            writer.WriteAttributeString("EnableSaveDataWithReport", .EnableSaveDataWithReport)
            writer.WriteAttributeString("EnableSavePreviewPicture", .EnableSavePreviewPicture)
            writer.WriteAttributeString("EnableSaveSummariesWithReport", .EnableSaveSummariesWithReport)
            writer.WriteAttributeString("EnableUseDummyData", .EnableUseDummyData)
            writer.WriteAttributeString("InitialDataContext", .InitialDataContext)
            writer.WriteAttributeString("InitialReportPartName", .InitialReportPartName)
        End With
        writer.WriteEndElement()    '/ReportOptions

    End Sub

    Private Sub GetPrintOptions(ByRef report As ReportDocument, ByRef writer As XmlWriter)

        writer.WriteStartElement("PrintOptions")
        With report.PrintOptions

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

    Private Sub GetSubreports(ByRef report As ReportDocument, ByRef writer As XmlWriter)

        writer.WriteStartElement("Reports")
        For Each subreport As ReportDocument In report.Subreports

            'recursively process each ReportDocument
            ProcessReport(subreport, writer)

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

    Private Sub GetDatabase(ByRef report As ReportDocument, ByRef writer As XmlWriter)

        writer.WriteStartElement("Database")

        GetTableLinks(report, writer)
        GetTables(report, writer)

        writer.WriteEndElement()    '/Database

    End Sub

    Private Sub GetTableLinks(ByRef report As ReportDocument, ByRef writer As XmlWriter)

        writer.WriteStartElement("TableLinks")

        For Each TL As TableLink In report.Database.Links

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

    Private Sub GetTables(ByRef report As ReportDocument, ByRef writer As XmlWriter)

        writer.WriteStartElement("Tables")

        For Each T As Table In report.Database.Tables
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

    Private Sub GetDataDefinition(ByRef report As ReportDocument, ByRef writer As XmlWriter)

        writer.WriteStartElement("DataDefinition")
        With report.DataDefinition

            writer.WriteElementString("GroupSelectionFormula", .GroupSelectionFormula)

            writer.WriteElementString("RecordSelectionFormula", .RecordSelectionFormula)

            writer.WriteStartElement("Groups")
            For Each group As Group In .Groups
                writer.WriteStartElement("Group")
                With group
                    writer.WriteAttributeString("ConditionField", .ConditionField.FormulaName)
                    'writer.WriteAttributeString("GroupOptions", .GroupOptions.Condition.)
                End With
                writer.WriteEndElement()    '/Group
            Next
            writer.WriteEndElement()    '/Groups

            writer.WriteStartElement("SortFields")
            For Each sortField As SortField In .SortFields
                writer.WriteStartElement("SortField")
                With sortField
                    writer.WriteAttributeString("Field", .Field.FormulaName)
                    writer.WriteAttributeString("SortDirection", .SortDirection.ToString)
                    writer.WriteAttributeString("SortType", .SortType.ToString)
                End With
                writer.WriteEndElement()    '/SortField
            Next
            writer.WriteEndElement()    '/SortFields

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

            writer.WriteStartElement("ParameterFieldDefinitions")
            For Each field In .ParameterFields
                GetFieldObject(field, writer)
            Next
            writer.WriteEndElement()

            writer.WriteStartElement("RunningTotalFieldDefinitions")
            For Each field In .RunningTotalFields
                GetFieldObject(field, writer)
            Next
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
                writer.WriteAttributeString("UseCount", .UseCount)
                writer.WriteAttributeString("ValueType", .ValueType.ToString)
                writer.WriteString(.Text)
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

#Region "Formatting"

    Private Sub GetAreaFormat(ByVal areaFormat As AreaFormat, ByRef writer As XmlWriter)

        writer.WriteStartElement("AreaFormat")
        With areaFormat

            writer.WriteAttributeString("EnableHideForDrillDown", .EnableHideForDrillDown)
            writer.WriteAttributeString("EnableKeepTogether", .EnableKeepTogether)
            writer.WriteAttributeString("EnableNewPageAfter", .EnableNewPageAfter)
            writer.WriteAttributeString("EnableNewPageBefore", .EnableNewPageBefore)
            writer.WriteAttributeString("EnablePrintAtBottomOfPage", .EnablePrintAtBottomOfPage)
            writer.WriteAttributeString("EnableResetPageNumberAfter", .EnableResetPageNumberAfter)
            writer.WriteAttributeString("EnableSuppress", .EnableSuppress)

        End With
        writer.WriteEndElement()    '/AreaFormat

    End Sub

    Private Sub GetBorderFormat(ByVal border As Border, ByRef writer As XmlWriter)

        writer.WriteStartElement("Border")
        With border
            writer.WriteAttributeString("BottomLineStyle", .BottomLineStyle.ToString)
            writer.WriteAttributeString("HasDropShadow", .HasDropShadow)
            writer.WriteAttributeString("LeftLineStyle", .LeftLineStyle.ToString)
            writer.WriteAttributeString("RightLineStyle", .RightLineStyle.ToString)
            writer.WriteAttributeString("TopLineStyle", .TopLineStyle.ToString)
            If (Me.ShowFormatTypes And FormatTypes.Color) = FormatTypes.Color Then GetColorFormat(.BackgroundColor, writer, "BackgroundColor")
            If (Me.ShowFormatTypes And FormatTypes.Color) = FormatTypes.Color Then GetColorFormat(.BorderColor, writer, "BorderColor")
        End With
        writer.WriteEndElement()    '/Border

    End Sub

    Private Sub GetColorFormat(ByVal color As Drawing.Color, ByRef writer As XmlWriter, Optional ByVal elementName As String = "Color")

        writer.WriteStartElement(elementName)
        With color
            writer.WriteAttributeString("Name", .Name)
            writer.WriteAttributeString("A", .A)
            writer.WriteAttributeString("R", .R)
            writer.WriteAttributeString("G", .G)
            writer.WriteAttributeString("B", .B)
        End With
        writer.WriteEndElement()    '/elementName

    End Sub

    Private Sub GetFontFormat(ByVal font As Drawing.Font, ByRef writer As XmlWriter)

        writer.WriteStartElement("Font")
        With font
            writer.WriteAttributeString("Bold", .Bold)
            writer.WriteAttributeString("FontFamily", .FontFamily.Name)
            writer.WriteAttributeString("GdiCharSet", .GdiCharSet)
            writer.WriteAttributeString("GdiVerticalFont", .GdiVerticalFont)
            writer.WriteAttributeString("Height", .Height)
            writer.WriteAttributeString("IsSystemFont", .IsSystemFont)
            writer.WriteAttributeString("Italic", .Italic)
            writer.WriteAttributeString("Name", .Name)
            writer.WriteAttributeString("OriginalFontName", .OriginalFontName)
            writer.WriteAttributeString("Size", .Size)
            writer.WriteAttributeString("SizeInPoints", .SizeInPoints)
            writer.WriteAttributeString("Strikeout", .Strikeout)
            writer.WriteAttributeString("Style", .Style)
            writer.WriteAttributeString("SystemFontName", .SystemFontName)
            writer.WriteAttributeString("Underline", .Underline)
            writer.WriteAttributeString("Unit", .Unit)
        End With
        writer.WriteEndElement()    '/Font

    End Sub

    Private Sub GetObjectFormat(ByVal objectFormat As ObjectFormat, ByRef writer As XmlWriter)

        writer.WriteStartElement("ObjectFormat")
        With objectFormat
            writer.WriteAttributeString("CssClass", .CssClass)
            writer.WriteAttributeString("EnableCanGrow", .EnableCanGrow)
            writer.WriteAttributeString("EnableCloseAtPageBreak", .EnableCloseAtPageBreak)
            writer.WriteAttributeString("EnableKeepTogether", .EnableKeepTogether)
            writer.WriteAttributeString("EnableSuppress", .EnableSuppress)
            writer.WriteAttributeString("HorizontalAlignment", .HorizontalAlignment.ToString)
        End With
        writer.WriteEndElement()    '/ObjectFormat

    End Sub

    Private Sub GetSectionFormat(ByVal sectionFormat As SectionFormat, ByRef writer As XmlWriter)

        writer.WriteStartElement("SectionFormat")
        With sectionFormat
            writer.WriteAttributeString("CssClass", .CssClass)
            writer.WriteAttributeString("EnableKeepTogether", .EnableKeepTogether)
            writer.WriteAttributeString("EnableNewPageAfter", .EnableNewPageAfter)
            writer.WriteAttributeString("EnableNewPageBefore", .EnableNewPageBefore)
            writer.WriteAttributeString("EnablePrintAtBottomOfPage", .EnablePrintAtBottomOfPage)
            writer.WriteAttributeString("EnableResetPageNumberAfter", .EnableResetPageNumberAfter)
            writer.WriteAttributeString("EnableSuppress", .EnableSuppress)
            writer.WriteAttributeString("EnableSuppressIfBlank", .EnableSuppressIfBlank)
            writer.WriteAttributeString("EnableUnderlaySection", .EnableUnderlaySection)
            If (Me.ShowFormatTypes And FormatTypes.Color) = FormatTypes.Color Then GetColorFormat(.BackgroundColor, writer, "BackgroundColor")
        End With
        writer.WriteEndElement()    '/SectionFormat

    End Sub

#End Region

#Region "Report Definition"

    Private Sub GetReportDefinition(ByRef report As ReportDocument, ByRef writer As XmlWriter)

        writer.WriteStartElement("ReportDefinition")

        GetAreas(report.ReportDefinition, writer)

        writer.WriteEndElement()    '/ReportDefinition

    End Sub

    Private Sub GetAreas(ByRef reportDefinition As ReportDefinition, ByRef writer As XmlWriter)

        writer.WriteStartElement("Areas")

        For Each area As Area In reportDefinition.Areas

            writer.WriteStartElement("Area")
            With area

                writer.WriteAttributeString("Kind", .Kind.ToString)
                writer.WriteAttributeString("Name", .Name)

                If (Me.ShowFormatTypes And FormatTypes.AreaFormat) = FormatTypes.AreaFormat Then GetAreaFormat(area.AreaFormat, writer)

                GetSections(area, writer)

            End With
            writer.WriteEndElement()    '/Area

        Next

        writer.WriteEndElement()    '/Areas

    End Sub

    Private Sub GetSections(ByRef area As Area, ByRef writer As XmlWriter)

        writer.WriteStartElement("Sections")

        For Each section As Section In area.Sections

            writer.WriteStartElement("Section")
            With section

                writer.WriteAttributeString("Height", .Height)
                writer.WriteAttributeString("Kind", .Kind.ToString)
                writer.WriteAttributeString("Name", .Name)

                If (Me.ShowFormatTypes And FormatTypes.SectionFormat) = FormatTypes.SectionFormat Then GetSectionFormat(section.SectionFormat, writer)

                GetReportObjects(section, writer)

            End With
            writer.WriteEndElement()    '/Section

        Next

        writer.WriteEndElement()    '/Sections

    End Sub

    Private Sub GetReportObjects(ByRef section As Section, ByRef writer As XmlWriter)

        writer.WriteStartElement("ReportObjects")

        For Each reportObject As ReportObject In section.ReportObjects

            writer.WriteStartElement(TypeName(reportObject))
            With reportObject

                'generic attributes
                writer.WriteAttributeString("Name", .Name)
                writer.WriteAttributeString("Kind", .Kind.ToString)

                writer.WriteAttributeString("Top", .Top)
                writer.WriteAttributeString("Left", .Left)
                writer.WriteAttributeString("Width", .Width)
                writer.WriteAttributeString("Height", .Height)

                If TypeOf reportObject Is BoxObject Then
                    With CType(reportObject, BoxObject)
                        writer.WriteAttributeString("Bottom", .Bottom)
                        writer.WriteAttributeString("EnableExtendToBottomOfSection", .EnableExtendToBottomOfSection)
                        writer.WriteAttributeString("EndSectionName", .EndSectionName)
                        writer.WriteAttributeString("LineStyle", .LineStyle.ToString)
                        writer.WriteAttributeString("LineThickness", .LineThickness)
                        writer.WriteAttributeString("Right", .Right)
                        If (Me.ShowFormatTypes And FormatTypes.Color) = FormatTypes.Color Then GetColorFormat(.LineColor, writer, "LineColor")
                    End With

                ElseIf TypeOf reportObject Is BlobFieldObject Then
                    'no specific properties
                ElseIf TypeOf reportObject Is ChartObject Then
                    'no specific properties
                ElseIf TypeOf reportObject Is CrossTabObject Then
                    'no specific properties
                ElseIf TypeOf reportObject Is DrawingObject Then
                    With CType(reportObject, DrawingObject)
                        writer.WriteAttributeString("Bottom", .Bottom)
                        writer.WriteAttributeString("EnableExtendToBottomOfSection", .EnableExtendToBottomOfSection)
                        writer.WriteAttributeString("EndSectionName", .EndSectionName)
                        writer.WriteAttributeString("LineStyle", .LineStyle.ToString)
                        writer.WriteAttributeString("LineThickness", .LineThickness)
                        writer.WriteAttributeString("Right", .Right)
                        If (Me.ShowFormatTypes And FormatTypes.Color) = FormatTypes.Color Then GetColorFormat(.LineColor, writer, "LineColor")
                    End With

                ElseIf TypeOf reportObject Is FieldHeadingObject Then
                    With CType(reportObject, FieldHeadingObject)
                        writer.WriteAttributeString("FieldObjectName", .FieldObjectName)
                        writer.WriteElementString("Text", .Text)
                    End With

                ElseIf TypeOf reportObject Is FieldObject Then
                    With CType(reportObject, FieldObject)
                        'writer.WriteAttributeString("FieldFormat", .FieldFormat)
                        writer.WriteAttributeString("DataSource", .DataSource.FormulaName)
                        'If (Me.ShowObjectTypes And ObjectTypes.ReportObject) = ObjectTypes.ReportObject Then
                        If (Me.ShowFormatTypes And FormatTypes.Color) = FormatTypes.Color Then GetColorFormat(.Color, writer)
                        If (Me.ShowFormatTypes And FormatTypes.Font) = FormatTypes.Font Then GetFontFormat(.Font, writer)
                        'End If

                    End With

                ElseIf TypeOf reportObject Is GraphicObject Then
                    'no specific properties
                ElseIf TypeOf reportObject Is LineObject Then
                    With CType(reportObject, LineObject)
                        writer.WriteAttributeString("Bottom", .Bottom)
                        writer.WriteAttributeString("EnableExtendToBottomOfSection", .EnableExtendToBottomOfSection)
                        writer.WriteAttributeString("EndSectionName", .EndSectionName)
                        writer.WriteAttributeString("LineStyle", .LineStyle.ToString)
                        writer.WriteAttributeString("LineThickness", .LineThickness)
                        writer.WriteAttributeString("Right", .Right)
                        If (Me.ShowFormatTypes And FormatTypes.Color) = FormatTypes.Color Then GetColorFormat(.LineColor, writer, "LineColor")
                    End With

                ElseIf TypeOf reportObject Is MapObject Then
                    'no specific properties
                ElseIf TypeOf reportObject Is OlapGridObject Then
                    'no specific properties
                ElseIf TypeOf reportObject Is PictureObject Then
                    'no specific properties
                ElseIf TypeOf reportObject Is SubreportObject Then

                ElseIf TypeOf reportObject Is TextObject Then
                    With CType(reportObject, TextObject)
                        writer.WriteElementString("Text", .Text)
                        ' If (Me.ShowObjectTypes And ObjectTypes.ReportObject) = ObjectTypes.ReportObject Then
                        If (Me.ShowFormatTypes And FormatTypes.Color) = FormatTypes.Color Then GetColorFormat(.Color, writer)
                        If (Me.ShowFormatTypes And FormatTypes.Font) = FormatTypes.Font Then GetFontFormat(.Font, writer)
                        'End If

                    End With

                End If

                'generic elements
                'If (Me.ShowObjectTypes And ObjectTypes.ReportObject) = ObjectTypes.ReportObject Then
                If (Me.ShowFormatTypes And FormatTypes.Border) = FormatTypes.Border Then GetBorderFormat(.Border, writer)
                If (Me.ShowFormatTypes And FormatTypes.ObjectFormat) = FormatTypes.ObjectFormat Then GetObjectFormat(.ObjectFormat, writer)
                'End If

            End With
            writer.WriteEndElement()    '/ReportObject

        Next

        writer.WriteEndElement()    '/ReportObjects

    End Sub

#End Region

End Class