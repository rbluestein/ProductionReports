Public Class Rpt_BVIProduction

    ' //
    ' // This report is intentionally set to update the spreadsheet for the current year only.
    ' // If an update for a previous year is required, the report date must be manually overridden.
    ' //

#Region " Declarations "
    Public Event NotifyForm(ByRef NotifyFormArgs As NotifyFormArgs)
    Public e As New NotifyFormArgs(NotifyFormArgs.SourceEnum.Rpt_BVIProduction)
    Private cEnviro As Enviro
    Private cCommon As New Common
    Private cReport As Report
    Private cReportConfig As New BVIProductionReportConfig
    Private cExcel_Generic As New Excel_Generic(cReportConfig)
    Private cActiveExcel As Microsoft.Office.Interop.Excel.Application
    Private cOutputFullPath As String
#End Region

    Public Sub New()
        cEnviro = gEnviro
    End Sub

    Public ReadOnly Property OutputFullPath() As String
        Get
            Return cOutputFullPath
        End Get
    End Property

    Public Function Init(ByVal OverrideReportDate As String) As Results
        Dim i As Integer
        Dim MyResults As New Results
        Dim ExcelResults As Results
        Dim SpreadsheetColl As Collection
        Dim LastReportDate As Date
        'Dim ClientItem As ClientItem
        Dim PartialYearItem As PartialYearItem
        Dim TotalsItem As TotalsItem
        Dim SectionOneColl As Collection
        Dim SectionTwoDailyColl As Collection
        Dim SectionTwoTotalsColl As Collection
        Dim RangeColl As Collection
        Dim SheetNum As Integer
        Dim ClientProductColl As Collection
        Dim ExcelAddress As New Excel_Generic.ExcelAddress
        Dim StartRow As Integer
        Dim ReportYear As String
        Dim Value As String
        Dim ThisDate As Date
        Dim ThisRow As Integer
        Dim FirstDate As Date
        Dim UseModifiedStartDateInd As Boolean = False
        Dim NewStartDate As DateTime = "01/01/2014"

        Try

            ' ___ Output
            e.Message = "Init"
            RaiseEvent NotifyForm(e)

            ' ___ Report date
            'If OverrideReportDate = Nothing Then
            '    LastReportDate = cReportConfig.ReportDate
            '    If LastReportDate.DayOfWeek = DayOfWeek.Monday Then
            '        LastReportDate = LastReportDate.AddDays(-3)
            '    Else
            '        LastReportDate = LastReportDate.AddDays(-1)
            '    End If
            'Else
            '    LastReportDate = OverrideReportDate
            'End If



            LastReportDate = OverrideReportDate


            ReportYear = LastReportDate.ToString("yyyy")


            ' ____ Get the address of the first row in the range
            SpreadsheetColl = cExcel_Generic.GetRangeData("_" & ReportYear & "_Dates")
            FirstDate = SpreadsheetColl("FirstValue")
            StartRow = SpreadsheetColl("xlAddr1RowNum")

            If ReportYear = "2009" Then
                ThisDate = cReportConfig.ReportStartDate
            Else
                ' ThisDate = "1/1/" & ReportYear
                ThisDate = FirstDate
            End If


            ' //
            ' // CLIENT DAILY VALUES
            ' //

            SectionOneColl = GetSectionOneColl(ReportYear)

            If 0 = 0 Then
                Do

                    DailyOne(UseModifiedStartDateInd, NewStartDate, SpreadsheetColl, ReportYear, ThisDate, SectionOneColl, ClientProductColl, ExcelAddress, ExcelResults, StartRow, SheetNum, ThisRow)

                    '    SpreadsheetColl = cExcel_Generic.DateLookup("_" & ReportYear & "_Dates", ThisDate.ToString("MM/dd/yyyy"))
                    '    SheetNum = SpreadsheetColl("SheetNum")
                    '    ThisRow = StartRow + SpreadsheetColl("RowOffset")
                    '    cActiveExcel = SpreadsheetColl("ActiveExcel")

                    '    ExcelAddress.SheetNum = SheetNum
                    '    ExcelAddress.RowNum = ThisRow

                    '    For i = 1 To SectionOneColl.Count
                    '        ClientItem = SectionOneColl(i)
                    '        If (Not IsDBNull(ClientItem.ClientStartDate)) AndAlso cCommon.IsDateBetween(ClientItem.ClientStartDate, ClientItem.ClientEndDate, ThisDate) Then
                    '            ClientProductColl = GetClientItemValues(ThisDate, ClientItem)
                    '            SpreadsheetColl = cExcel_Generic.GetRangeData("_" & ReportYear & "_" & ClientItem.ClientID)
                    '            ExcelAddress.ColumnLtr = SpreadsheetColl("xlAddr1ColumnName")
                    '            ExcelResults = cExcel_Generic.ExportCollectionToExcel(ClientProductColl, ExcelAddress, cActiveExcel)
                    '            e.Message = "Rpt_BVIProduction.Init: Processing " & ThisDate.ToString("MM/dd/yyyy") & " " & ClientItem.ClientID
                    '            RaiseEvent NotifyForm(e)
                    '            cActiveExcel = ExcelResults.Value
                    '        End If
                    '    Next

                    '    ThisDate = ThisDate.AddDays(1)

                    '    ' ___ Output
                    '    e.Message = "Rpt_BVIProdction.Init ThisDate: " & ThisDate
                    '    RaiseEvent NotifyForm(e)

                Loop Until cCommon.DateCompare(ThisDate, LastReportDate, True) = 1

                ' ___ Display report run date
                ExcelAddress = New Excel_Generic.ExcelAddress
                ExcelAddress.RangeName = "_" & ReportYear & "_ReportRunDate"
                cExcel_Generic.ExportFieldToExcel(cCommon.GetServerDateTime.ToString("MMM-dd-yyyy"), ExcelAddress, cActiveExcel)

                cActiveExcel = cExcel_Generic.SetFill(SheetNum, ThisRow)

            End If

            ' //
            ' // CUMULATIVE VALUES
            ' //

            ' ___ Standard cum items
            ExcelAddress.RangeName = Nothing
            SectionTwoDailyColl = GetSectionTwoDailyColl(ReportYear)
            For i = 1 To SectionTwoDailyColl.Count
                PartialYearItem = SectionTwoDailyColl(i)

                ' ___ Get the row number
                'SpreadsheetColl = cExcel_Generic.GetRangeData("_" & ReportYear & "_" & ClientItem.ProductID)
                SpreadsheetColl = cExcel_Generic.GetRangeData(PartialYearItem.RangeName)
                cActiveExcel = SpreadsheetColl("ActiveExcel")
                SheetNum = SpreadsheetColl("SheetNum")
                ThisRow = SpreadsheetColl("xlAddr1RowNum")
                ExcelAddress.SheetNum = SheetNum
                ExcelAddress.RowNum = ThisRow

                ' ___ Get the column number
                'SpreadsheetColl = cExcel_Generic.GetRangeData("_" & ReportYear & "_" & ClientItem.ClientID)
                SpreadsheetColl = cExcel_Generic.GetRangeData("_" & ReportYear & "_" & PartialYearItem.ClientID)
                cActiveExcel = SpreadsheetColl("ActiveExcel")
                ExcelAddress.ColumnLtr = SpreadsheetColl("xlAddr1ColumnName")

                ' ___ Get the value
                Value = GetCumValue(PartialYearItem, SectionOneColl(PartialYearItem.ClientID), FirstDate, LastReportDate)
                ExcelResults = cExcel_Generic.ExportFieldToExcel(Value, ExcelAddress, cActiveExcel)
                e.Message = "Rpt_BVIProduction.Init:  Processing cumulative values ClientID: " & PartialYearItem.ClientID & ", Product: " & PartialYearItem.ProductID
                RaiseEvent NotifyForm(e)
                cActiveExcel = ExcelResults.Value
            Next


            ' ___ Multi-year cum items
            SectionTwoTotalsColl = GetSectionTwoTotalsColl(ReportYear)
            For i = 1 To SectionTwoTotalsColl.Count
                TotalsItem = SectionTwoTotalsColl(i)

                ' ___ Get the row number
                SpreadsheetColl = cExcel_Generic.GetRangeData(TotalsItem.RangeName)
                cActiveExcel = SpreadsheetColl("ActiveExcel")
                SheetNum = SpreadsheetColl("SheetNum")
                ThisRow = SpreadsheetColl("xlAddr1RowNum")
                ExcelAddress.SheetNum = SheetNum
                ExcelAddress.RowNum = ThisRow

                ' ___ Get the column number
                If TotalsItem.ColumnNameOverride = Nothing Then
                    SpreadsheetColl = cExcel_Generic.GetRangeData("_" & TotalsItem.ExternalValueYear & "_" & TotalsItem.ClientID)
                Else
                    SpreadsheetColl = cExcel_Generic.GetRangeData("_" & TotalsItem.ExternalValueYear & "_" & TotalsItem.ColumnNameOverride)
                End If

                cActiveExcel = SpreadsheetColl("ActiveExcel")
                ExcelAddress.ColumnLtr = SpreadsheetColl("xlAddr1ColumnName")

                ' ___ Get the value
                Value = TotalsItem.Value
                ExcelResults = cExcel_Generic.ExportFieldToExcel(Value, ExcelAddress, cActiveExcel)
                e.Message = "Rpt_BVIProduction.Init:  Processing multi-year cumulative values ClientID: " & TotalsItem.ClientID & ", Product: " & TotalsItem.ProductID
                RaiseEvent NotifyForm(e)
                cActiveExcel = ExcelResults.Value
            Next

            ' ___ Range items
            RangeColl = GetRangeColl(ReportYear)
            If RangeColl.Count > 0 Then
                'Dim Dec01 As Integer
                'Dec01 = cExcel_Generic.GetDateRow("_" & ReportYear & "_Dates", "12/01/2010")

                For i = 1 To RangeColl.Count
                    SpreadsheetColl = cExcel_Generic.GetRangeData(RangeColl(i).RangeName)
                    cActiveExcel = SpreadsheetColl("ActiveExcel")
                    SheetNum = SpreadsheetColl("SheetNum")
                    ThisRow = SpreadsheetColl("xlAddr1RowNum")
                    ExcelAddress.ColumnLtr = SpreadsheetColl("xlAddr1ColumnName")
                    ExcelAddress.SheetNum = SheetNum
                    ExcelAddress.RowNum = ThisRow
                    Value = RangeColl(i).Value
                    ExcelResults = cExcel_Generic.ExportFieldToExcel(Value, ExcelAddress, cActiveExcel)
                    cActiveExcel = ExcelResults.Value

                    'ExcelAddress.RowNum = Dec01
                    ExcelResults = cExcel_Generic.ExportFieldToExcelForSum(Value, ExcelAddress, cActiveExcel)
                    cActiveExcel = ExcelResults.Value

                Next
            End If

            cExcel_Generic.Finish(cActiveExcel)

            ' ___ OutputFullPath
            cOutputFullPath = cExcel_Generic.OutputFullPath

            MyResults.Success = True
            MyResults.Value = LastReportDate
            Return MyResults

        Catch ex As Exception
            cReport = New Report
            cReport.Report("Error# 1450: Rpt_BVIProduction.Init " & ex.Message, Report.ReportTypeEnum.LogError)
        End Try
    End Function

    Private Sub DailyOne(ByVal UseModifiedStartDateInd As Boolean, ByVal NewStartDate As DateTime, ByRef SpreadsheetColl As Collection, ByVal ReportYear As String, ByRef ThisDate As DateTime, ByRef SectionOneColl As Collection, ByRef ClientProductColl As Collection, ByRef ExcelAddress As Excel_Generic.ExcelAddress, ByRef ExcelResults As Results, ByVal StartRow As Integer, ByRef SheetNum As Integer, ByRef ThisRow As Integer)
        Try
            If UseModifiedStartDateInd Then
                If NewStartDate >= ThisDate Then
                    SpreadsheetColl = cExcel_Generic.DateLookup("_" & ReportYear & "_Dates", ThisDate.ToString("MM/dd/yyyy"))
                    ThisDate = ThisDate.AddDays(1)
                    e.Message = "Rpt_BVIProdction.Init ThisDate: " & ThisDate
                    RaiseEvent NotifyForm(e)
                Else
                    DailyTwo(UseModifiedStartDateInd, NewStartDate, SpreadsheetColl, ReportYear, ThisDate, SectionOneColl, ClientProductColl, ExcelAddress, ExcelResults, StartRow, SheetNum, ThisRow)
                End If
            Else
                DailyTwo(UseModifiedStartDateInd, NewStartDate, SpreadsheetColl, ReportYear, ThisDate, SectionOneColl, ClientProductColl, ExcelAddress, ExcelResults, StartRow, SheetNum, ThisRow)
            End If
        Catch ex As Exception
            cReport = New Report
            cReport.Report("Error# 1454: Rpt_BVIProduction.DailyOne " & ex.Message, Report.ReportTypeEnum.LogError)
        End Try
    End Sub

    Private Sub DailyTwo(ByVal UseModifiedStartDateInd As Boolean, ByVal NewStartDate As DateTime, ByRef SpreadsheetColl As Collection, ByVal ReportYear As String, ByRef ThisDate As DateTime, ByRef SectionOneColl As Collection, ByRef ClientProductColl As Collection, ByRef ExcelAddress As Excel_Generic.ExcelAddress, ByRef ExcelResults As Results, ByVal StartRow As Integer, ByRef SheetNum As Integer, ByRef ThisRow As Integer)
        Dim i As Integer
        Dim ClientItem As ClientItem

        Try
            SpreadsheetColl = cExcel_Generic.DateLookup("_" & ReportYear & "_Dates", ThisDate.ToString("MM/dd/yyyy"))
            SheetNum = SpreadsheetColl("SheetNum")
            ThisRow = StartRow + SpreadsheetColl("RowOffset")
            cActiveExcel = SpreadsheetColl("ActiveExcel")

            ExcelAddress.SheetNum = SheetNum
            ExcelAddress.RowNum = ThisRow

            For i = 1 To SectionOneColl.Count
                ClientItem = SectionOneColl(i)
                If (Not IsDBNull(ClientItem.ClientStartDate)) AndAlso cCommon.IsDateBetween(ClientItem.ClientStartDate, ClientItem.ClientEndDate, ThisDate) Then
                    ClientProductColl = GetClientItemValues(ThisDate, ClientItem)
                    SpreadsheetColl = cExcel_Generic.GetRangeData("_" & ReportYear & "_" & ClientItem.ClientID)
                    ExcelAddress.ColumnLtr = SpreadsheetColl("xlAddr1ColumnName")
                    ExcelResults = cExcel_Generic.ExportCollectionToExcel(ClientProductColl, ExcelAddress, cActiveExcel)
                    e.Message = "Rpt_BVIProduction.Init: Processing " & ThisDate.ToString("MM/dd/yyyy") & " " & ClientItem.ClientID
                    RaiseEvent NotifyForm(e)
                    cActiveExcel = ExcelResults.Value
                End If
            Next

            ThisDate = ThisDate.AddDays(1)

            ' ___ Output
            e.Message = "Rpt_BVIProdction.Init ThisDate: " & ThisDate
            RaiseEvent NotifyForm(e)

        Catch ex As Exception
            cReport = New Report
            cReport.Report("Error# 1453: Rpt_BVIProduction.DailyTwo " & ex.Message, Report.ReportTypeEnum.LogError)
        End Try        
    End Sub


    Private Function GetCumValue(ByRef PartialYearItem As PartialYearItem, ByVal ClientItem As ClientItem, ByVal FirstReportDate As Date, ByVal ReportDate As Date) As Decimal
        Dim Sql As New System.Text.StringBuilder
        Dim Querypack As QueryPack
        Dim Results As Decimal
        Dim ReportStartDate As String
        Dim ReportEndDate As String
        Dim QueryStartDate As String
        Dim QueryEndDate As String
        Dim DateEligibleInd As Boolean

        Try

            ' ___ Report start and end dates
            If ReportDate.ToString("yyyy") = "2009" Then
                ReportStartDate = "8/12/2009"
            Else
                ReportStartDate = FirstReportDate.ToString("M/d/yyyy")
            End If
            ReportEndDate = ReportDate

            ' ___ Do any of this client's dates fall between the report start and end dates?

            ' ___ The client start date must fall on or before the report end date.
            If (Not IsDBNull(ClientItem.ClientStartDate)) AndAlso cCommon.DateCompare(ClientItem.ClientStartDate, ReportEndDate, True) < 1 Then

                ' ___ The client end date must not fall before the report startdate
                If IsDate(ClientItem.ClientEndDate) Then
                    If cCommon.DateCompare(ClientItem.ClientEndDate, ReportStartDate, True) > -1 Then
                        DateEligibleInd = True
                    End If
                Else
                    DateEligibleInd = True
                End If

            End If

            If DateEligibleInd Then

                ' ___ Narrow the query dates to reflect client dates, as required.

                ' ___ If the client start date falls after the report start date, set the query start date to the client start date.
                If cCommon.DateCompare(ClientItem.ClientStartDate, ReportStartDate, True) = 1 Then
                    QueryStartDate = ClientItem.ClientStartDate
                Else
                    QueryStartDate = ReportStartDate
                End If

                ' ___ If the client end date falls before the report end date, set the query end date to the client end date.
                If IsDate(ClientItem.ClientEndDate) Then
                    If cCommon.DateCompare(ClientItem.ClientEndDate, ReportEndDate, True) = -1 Then
                        QueryEndDate = ClientItem.ClientEndDate
                    Else
                        QueryEndDate = ReportEndDate
                    End If
                Else
                    QueryEndDate = ReportEndDate
                End If



                ' ___ Execute the query
                Sql.Append("SELECT ")
                Sql.Append("IsNull(SUM(Cast(IsNull(FieldData, 0) as decimal(10,2))), 0) ")
                Sql.Append("FROM Rpt_ProductHistory ")
                Sql.Append("WHERE ClientID = '" & PartialYearItem.ClientID & "' AND ProductID = '" & PartialYearItem.ProductID & "' AND ")
                Sql.Append("dbo.ufn_IsDateBetween(CallDate, '" & QueryStartDate & "', '" & QueryEndDate & "') = 1 AND ")
                Sql.Append("FieldName = 'AnnualPremium'")

                Querypack = cCommon.GetDTWithQuerypack(Sql.ToString)
                If Not Querypack.Success Then
                    cReport.Report("Error# 1451a: Rpt_EnrollCenterMonthly.GetCumValue " & Querypack.TechErrMsg, Report.ReportTypeEnum.LogError)
                End If

                Results = Querypack.dt.rows(0)(0)
            End If

            ' ___ Apply the external cumulative values
            If ReportDate.ToString("yyyy") = PartialYearItem.ExternalValueYear Then
                Results += PartialYearItem.Value
            End If

            Return Results

        Catch ex As Exception
            cReport = New Report
            cReport.Report("Error# 1451d: Rpt_BVIProduction.GetCumValue " & ex.Message, Report.ReportTypeEnum.LogError)
        End Try
    End Function

    Private Function GetClientItemValues(ByVal ReportDate As Date, ByVal ClientItem As ClientItem) As Collection
        Dim i, j As Integer
        Dim Box() As String
        Dim Box2() As String
        Dim Sql As New System.Text.StringBuilder
        Dim Querypack As QueryPack
        Dim Coll As New Collection
        Dim BlendedValue As Decimal
        Dim CompositeClientInd As Boolean
        Dim ProductID As String
        Dim SubClientID As String
        Dim SubProductID As String

        Try

            ' // Break out into separate column 1/1/2010

            If InStr(ClientItem.ClientID, "|") > 0 Then
                CompositeClientInd = True
            End If

            If CompositeClientInd Then

                ' ___ Composite client
                Box = Split(ClientItem.ProductID, "|")
                For i = 0 To Box.GetUpperBound(0)
                    SubClientID = Box(i).Substring(0, InStr(Box(i), "~") - 1)
                    SubProductID = cCommon.Right(Box(i), Box(i).Length - InStr(Box(i), "~"))
                    Sql.Length = 0
                    Sql.Append("SELECT ")
                    Sql.Append("IsNull(SUM(Cast(IsNull(FieldData, 0) as decimal(10,2))), 0) ")
                    Sql.Append("FROM Rpt_ProductHistory ")
                    Sql.Append("WHERE ClientID = '" & SubClientID & "' AND ProductID = '" & SubProductID & "' AND dbo.ufn_IsDateEqual(CallDate, '" & ReportDate & "') = 1 AND FieldName = 'AnnualPremium'")
                    Querypack = cCommon.GetDTWithQuerypack(Sql.ToString)
                    Coll.Add(Querypack.dt.rows(0)(0))
                Next

            Else

                ' ___ Standard Client
                Box = Split(ClientItem.ProductID, "|")
                For i = 0 To Box.GetUpperBound(0)
                    BlendedValue = 0
                    Box2 = Split(Box(i), "~")
                    For j = 0 To Box2.GetUpperBound(0)
                        ProductID = Box2(j)
                        Sql.Length = 0
                        Sql.Append("SELECT ")
                        Sql.Append("IsNull(SUM(Cast(IsNull(FieldData, 0) as decimal(10,2))), 0) ")
                        Sql.Append("FROM Rpt_ProductHistory ")
                        Sql.Append("WHERE ClientID = '" & ClientItem.ClientID & "' AND ProductID = '" & ProductID & "' AND dbo.ufn_IsDateEqual(CallDate, '" & ReportDate & "') = 1 AND FieldName = 'AnnualPremium'")
                        Querypack = cCommon.GetDTWithQuerypack(Sql.ToString)
                        BlendedValue += Querypack.dt.rows(0)(0)
                    Next
                    Coll.Add(BlendedValue)
                Next

            End If

            Return Coll

        Catch ex As Exception
            cReport = New Report
            cReport.Report("Error# 1452: Rpt_BVIProduction.GetClientProdColl " & ex.Message, Report.ReportTypeEnum.LogError)
        End Try
    End Function

#Region " Collections "
    Private Function GetSectionOneColl(ByVal ReportYear As String) As Collection
        Dim Coll As New Collection
        Dim dt As DataTable

        Try

            dt = cCommon.GetDT("SELECT ClusterID, StartDate, EndDate FROM Excel_Cluster WHERE ProdRptStatusInd = 1")

            ' Use ClientProductID
            Coll.Add(New ClientItem("Options|Choices", "OPTIONS~TMARKUL2|OPTIONS~TMARKCOMBO|CHOICES~TMARKUL2|CHOICES~TMARKCOMBO", dt), "Options|Choices")

            If ReportYear < "2011" Then
                Coll.Add(New ClientItem("HardRock", "TransDI|TransTerm|TMarkUL|TMarkCombo", dt), "HardRock")
                Coll.Add(New ClientItem("Genesis", "AllStateUL|AllStateCI", dt), "Genesis")
                Coll.Add(New ClientItem("Fulton", "BMWLife|BMDI|BMCCI", dt), "Fulton")
                Coll.Add(New ClientItem("COKC", "TransUL2~TRANSCANCERSELPLUS", dt), "COKC")
                Coll.Add(New ClientItem("HT", "BMWLife|AmGenCancer~AIGCCI|UnumAcc", dt), "HT")  ' spreadsheet says AIG-CI
            End If

            Coll.Add(New ClientItem("C3", "SunVSTD|TransUL|TransCCI|TransACC", dt), "C3")
            Coll.Add(New ClientItem("Morgans", "TransDI|TransTerm|TMarkUL|TMarkCombo", dt), "Morgans")
            Coll.Add(New ClientItem("Superior", "AllStateSTD|AllStateTerm|AllStateUL|AllStateCI", dt), "Superior")
            Coll.Add(New ClientItem("CTCA", "AllStateUL|AllStateCancer~AllStateCan|AllStateCI", dt), "CTCA")
            'Coll.Add(New ClientItem("Weathershield", "TMARKUL|TMARKCOMBO|TMARKACC|TMARKDI", dt), "Weathershield")


            Coll.Add(New ClientItem("Weathershield", "ALLSTATEDI|ALLSTATEUL|ALLSTATECI|ALLSTATECOMBO|ALLSTATEACC|TMARKUL|TMARKCOMBO|TMARKACC|TMARKDI", dt), "Weathershield")

            Coll.Add(New ClientItem("PTGaming", "AllStateDI|AllStateUL|AllStateCI|AllStateAcc|TmarkACC|TmarkCombo|TmarkDI|TmarkUL", dt), "PTGaming")
            Coll.Add(New ClientItem("Fortiss", "AllStateDI|AllStateUL|AllStateCI|AllStateAcc|TmarkACC|TmarkCombo|TmarkDI|TmarkUL", dt), "Fortiss")

            If ReportYear < 2011 Then
                Coll.Add(New ClientItem("BureauVeritas", "TransUL|AIGCCI|AIGACC", dt), "BureauVeritas")
            Else
                Coll.Add(New ClientItem("BureauVeritas", "AIGCCI|TRANSUL|AIGACC|TMARKUL~TMARKULNY|TMARKCOMBO|TMARKACC", dt), "BureauVeritas")
            End If

            If ReportYear < 2011 Then
                Coll.Add(New ClientItem("PKOH", "TMarkUL|TMarkCombo|TMarkAcc", dt), "PKOH")
            Else
                Coll.Add(New ClientItem("PKOH", "TMarkUL~TMARKULNY|TMarkCombo|TMarkAcc", dt), "PKOH")
            End If

            If ReportYear >= 2012 Then
                Coll.Add(New ClientItem("WisEdge", "TMarkUL|HumAcc", dt), "WISEDGE")
            End If

            If ReportYear >= 2013 Then
                Coll.Add(New ClientItem("MalibuMgt", "TMarkUL|TMarkCombo|TMarkAcc", dt), "MALIBUMGT")
                Coll.Add(New ClientItem("Neuterra", "TMarkUL|TMarkCombo|TMarkAcc", dt), "NEUTERRA")
                Coll.Add(New ClientItem("PACarp", "COMBINEDACC|COMBINEDDI|COMBINEDTERM|COMBINEDUL|FIDELTERM", dt), "PACarp")
                Coll.Add(New ClientItem("Cardinal", "TMarkUL", dt), "CARDINAL")
                Coll.Add(New ClientItem("Ease", "AllStateAcc|AllStateUL|AllStateHosp", dt), "EASE")
                Coll.Add(New ClientItem("Seasons", "LincAcc2|LincCCI|LincUL", dt), "SEASONS")
                Coll.Add(New ClientItem("Walsh", "TransAcc|TransCCI|TransDI|TransWL", dt), "WALSH")

                Coll.Add(New ClientItem("Gamestop", "TMarkUL", dt), "Gamestop")
            End If

            Coll.Add(New ClientItem("Martinrea", "TMarkUL|TMarkCombo|TMarkAcc", dt), "Martinrea")
            Coll.Add(New ClientItem("Stantec", "TransUL|TMarkUL|TMarkCombo", dt), "Stantec")
            Coll.Add(New ClientItem("IPBC", "TMarkUL|TMarkAcc|TransCCI|Legal", dt), "IPBC")
            Coll.Add(New ClientItem("Knighted", "TMarkUL|TMarkAcc|TMarkCombo", dt), "Knighted")
            Coll.Add(New ClientItem("GNC", "EyeMedVis3|AllStateUL|AllStateCI|AllStateAcc", dt), "GNC")
            Coll.Add(New ClientItem("DioPitt", "AllStateDI|AllStateUL|AllStateCI|AllStateAcc", dt), "DioPitt")
            Coll.Add(New ClientItem("RDS", "TransUL|TransAcc|TransCCI|Nothing", dt), "RDS")
            Coll.Add(New ClientItem("Armstrong", "TMarkCancer", dt), "Armstrong")
            Coll.Add(New ClientItem("SFWMD", "TMarkUL|TMarkCombo|TMarkAcc", dt), "SFWMD")

            If ReportYear >= 2011 Then
                Coll.Add(New ClientItem("Wolverine", "TransUL|TransAcc|TMarkUL|CombACC|LincDI|HUMVTR|HUMVTRSP|HUMVTRCH", dt), "Wolverine")
                Coll.Add(New ClientItem("ASCOM", "LincACC|LincCCI|LincUL", dt), "ASCOM")
                Coll.Add(New ClientItem("YMCA", "HUMCI|HUMACC|HUMWL", dt), "YMCA")
            End If

            Return Coll

        Catch ex As Exception
            cReport = New Report
            cReport.Report("Error# 1454: Rpt_BVIProduction.GetSectionOneColl " & ex.Message, Report.ReportTypeEnum.LogError)
        End Try
    End Function

    Private Function GetSectionTwoDailyColl(ByVal ReportYear As String) As Collection
        Dim Coll As New Collection

        Try
            Select Case ReportYear
                Case "2009"
                    Coll.Add(New PartialYearItem("RDS", "EyeMedVis", "_" & ReportYear & "_EyeMedVis", 86929.64, ReportYear))
                    Coll.Add(New PartialYearItem("BureauVeritas", "Legal", "_" & ReportYear & "_HyattLegal", 15048, ReportYear))
                    Coll.Add(New PartialYearItem("HT", "Legal", "_" & ReportYear & "_LegalClub", 23184, ReportYear))
                    Coll.Add(New PartialYearItem("PKOH", "Legal", "_" & ReportYear & "_LegalClub", 14112, ReportYear))
                    Coll.Add(New PartialYearItem("RDS", "Legal", "_" & ReportYear & "_LegalClub", 11453, ReportYear))
                    Coll.Add(New PartialYearItem("SFWMD", "Legal", "_" & ReportYear & "_LegalClub", 0, ReportYear))
                    Coll.Add(New PartialYearItem("Stantec", "Legal", "_" & ReportYear & "_LegalClub", 32924, ReportYear))

                Case "2010"
                    Coll.Add(New PartialYearItem("RDS", "EyeMedVis", "_" & ReportYear & "_EyeMedVis", 0, ReportYear))
                    Coll.Add(New PartialYearItem("BureauVeritas", "Legal", "_" & ReportYear & "_HyattLegal", 0, ReportYear))
                    Coll.Add(New PartialYearItem("HT", "Legal", "_" & ReportYear & "_LegalClub", 0, ReportYear))
                    Coll.Add(New PartialYearItem("PKOH", "Legal", "_" & ReportYear & "_LegalClub", 0, ReportYear))
                    Coll.Add(New PartialYearItem("RDS", "Legal", "_" & ReportYear & "_LegalClub", 0, ReportYear))
                    Coll.Add(New PartialYearItem("SFWMD", "Legal", "_" & ReportYear & "_LegalClub", 0, ReportYear))
                    Coll.Add(New PartialYearItem("Stantec", "Legal", "_" & ReportYear & "_LegalClub", 0, ReportYear))

                Case "2011"

                    ' ___ Discontinued 5/5/2011
                    'Coll.Add(New PartialYearItem("RDS", "EyeMedVis", "_" & ReportYear & "_EyeMedVis", 0, ReportYear))

                    Coll.Add(New PartialYearItem("BureauVeritas", "Legal", "_" & ReportYear & "_HyattLegal", 0, ReportYear))
                    Coll.Add(New PartialYearItem("PKOH", "Legal", "_" & ReportYear & "_LegalClub", 0, ReportYear))
                    Coll.Add(New PartialYearItem("RDS", "Legal", "_" & ReportYear & "_LegalClub", 0, ReportYear))
                    Coll.Add(New PartialYearItem("SFWMD", "Legal", "_" & ReportYear & "_LegalClub", 0, ReportYear))
                    Coll.Add(New PartialYearItem("Stantec", "Legal", "_" & ReportYear & "_LegalClub", 0, ReportYear))
                    Coll.Add(New PartialYearItem("Wolverine", "Legal", "_" & ReportYear & "_LegalClub", 0, ReportYear))

                Case "2012"
                    Coll.Add(New PartialYearItem("BureauVeritas", "Legal", "_" & ReportYear & "_HyattLegal", 0, ReportYear))
                    Coll.Add(New PartialYearItem("PKOH", "Legal", "_" & ReportYear & "_LegalClub", 0, ReportYear))
                    Coll.Add(New PartialYearItem("RDS", "Legal", "_" & ReportYear & "_LegalClub", 0, ReportYear))
                    Coll.Add(New PartialYearItem("SFWMD", "Legal", "_" & ReportYear & "_LegalClub", 0, ReportYear))
                    Coll.Add(New PartialYearItem("Stantec", "Legal", "_" & ReportYear & "_LegalClub", 0, ReportYear))
                    Coll.Add(New PartialYearItem("Wolverine", "Legal", "_" & ReportYear & "_LegalClub", 0, ReportYear))

                Case "2013", "2014", "2015"
                    Coll.Add(New PartialYearItem("BureauVeritas", "Legal", "_" & ReportYear & "_HyattLegal", 0, ReportYear))
                    Coll.Add(New PartialYearItem("Martinrea", "Legal", "_" & ReportYear & "_LegalShield", 0, ReportYear))
                    Coll.Add(New PartialYearItem("PKOH", "Legal", "_" & ReportYear & "_LegalClub", 0, ReportYear))
                    Coll.Add(New PartialYearItem("RDS", "Legal", "_" & ReportYear & "_LegalClub", 0, ReportYear))
                    Coll.Add(New PartialYearItem("SFWMD", "Legal", "_" & ReportYear & "_LegalClub", 0, ReportYear))
                    Coll.Add(New PartialYearItem("Stantec", "Legal", "_" & ReportYear & "_LegalClub", 0, ReportYear))
                    Coll.Add(New PartialYearItem("Wolverine", "Legal", "_" & ReportYear & "_LegalClub", 0, ReportYear))

            End Select

            Return Coll

        Catch ex As Exception
            cReport = New Report
            cReport.Report("Error# 1455: Rpt_BVIProduction.GetSectionTwoDailyColl " & ex.Message, Report.ReportTypeEnum.LogError)
        End Try
    End Function

    Private Function GetSectionTwoTotalsColl(ByVal ReportYear As String) As Collection
        Dim Coll As New Collection

        Try
            Select Case ReportYear
                Case "2009"
                    Coll.Add(New TotalsItem("RDS", "EZTrans", 16341, "2009", Nothing))
                    Coll.Add(New TotalsItem("Gentiva", "EZTrans", 23661, "2009", Nothing))
                    Coll.Add(New TotalsItem("LNT", "EZTrans", 42115, "2009", "LNT_EZ"))
                    Coll.Add(New TotalsItem("PepBoys", "EZTrans", 76061, "2009", "PepBoys_EZ"))
                    Coll.Add(New TotalsItem("Stantec", "EZTrans", 31745, "2009", Nothing))
                    Coll.Add(New TotalsItem("WirelessRetail", "EZTrans", 94, "2009", Nothing))
                    Coll.Add(New TotalsItem("BureauVeritas", "EZTrans", 20402, "2009", Nothing))

                    Coll.Add(New TotalsItem("AHL", "EZTmark", 208, "2009", Nothing))
                    Coll.Add(New TotalsItem("Amlings", "EZTmark", 52, "2009", Nothing))
                    Coll.Add(New TotalsItem("Aramark", "EZTmark", 1822, "2009", Nothing))
                    Coll.Add(New TotalsItem("Charming", "EZTmark", 1092, "2009", Nothing))
                    Coll.Add(New TotalsItem("Pennsylvania", "EZTmark", 156, "2009", Nothing))
                    Coll.Add(New TotalsItem("HardRock", "EZTmark", 1456, "2009", "HardRock_EZ"))
                    Coll.Add(New TotalsItem("RHSheppard", "EZTmark", 312, "2009", Nothing))
                    Coll.Add(New TotalsItem("RiteAid", "EZTmark", 416, "2009", Nothing))
                    Coll.Add(New TotalsItem("OptionsChoices", "EZTmark", 248618, "2009", Nothing))
                    Coll.Add(New TotalsItem("SFWMD", "EZTmark", 11962, "2009", Nothing))
                    Coll.Add(New TotalsItem("Vincents", "EZTmark", 52, "2009", Nothing))
                    Coll.Add(New TotalsItem("Telespectrum", "EZTmark", 52, "2009", Nothing))
                    Coll.Add(New TotalsItem("UFCW204", "EZTmark", 156, "2009", Nothing))

                    ' ___ Dec 29 2009 from Ron
                    Coll.Add(New TotalsItem("Superior", "EZAllstate", 0, "2009", "Superior_EZ"))
                    Coll.Add(New TotalsItem("CTCA", "EZAllstate", 884, "2009", Nothing))
                    Coll.Add(New TotalsItem("Weathershield", "EZAllstate", 14664, "2009", Nothing))
                    Coll.Add(New TotalsItem("Genesis", "EZAllstate", 104000, "2009", Nothing))


                Case "2010"

                    ' ___ Dec 3: Trustmark from Ron
                    Coll.Add(New TotalsItem("AHL", "EZTmark", 156, "2010", Nothing))
                    Coll.Add(New TotalsItem("Aramark", "EZTmark", 1613, "2010", Nothing))
                    Coll.Add(New TotalsItem("Charming", "EZTmark", 52, "2010", Nothing))
                    Coll.Add(New TotalsItem("Martinrea", "EZTmark", 20228, "2010", Nothing))
                    Coll.Add(New TotalsItem("Morgans", "EZTmark", 9672, "2010", Nothing))
                    Coll.Add(New TotalsItem("RHSheppard", "EZTmark", 312, "2010", Nothing))
                    Coll.Add(New TotalsItem("RiteAid", "EZTmark", 104, "2010", Nothing))
                    'Value for CCIU + SEIU combined with new info received Dec 9
                    'CCIU  $254,185.60 and SEIU  $287,990.00 
                    Coll.Add(New TotalsItem("OptionsChoices", "EZTmark", 542175.6, "2010", Nothing))
                    Coll.Add(New TotalsItem("SFWMD", "EZTmark", 13209, "2010", Nothing))
                    Coll.Add(New TotalsItem("Vincents", "EZTmark", 52, "2010", Nothing))
                    Coll.Add(New TotalsItem("UFCW204", "EZTmark", 104, "2010", Nothing))

                    ' ___ Dec 9: Allstate from Ron
                    Coll.Add(New TotalsItem("Weathershield", "EZAllstate", 12064, "2010", Nothing))
                    Coll.Add(New TotalsItem("CTCA", "EZAllstate", 832, "2010", Nothing))
                    Coll.Add(New TotalsItem("Genesis", "EZAllstate", 86008, "2010", Nothing))

                    ' ___ Dec 13: Transamerica from Ron
                    Coll.Add(New TotalsItem("RDS", "EZTrans", 25324, "2010", Nothing))
                    Coll.Add(New TotalsItem("LNT", "EZTrans", 3224, "2010", "LNT_EZ"))
                    Coll.Add(New TotalsItem("PepBoys", "EZTrans", 23088, "2010", "PepBoys_EZ"))
                    Coll.Add(New TotalsItem("Stantec", "EZTrans", 33592, "2010", Nothing))
                    Coll.Add(New TotalsItem("BureauVeritas", "EZTrans", 16588, "2010", Nothing))


                Case "2011"
                    ' ___ Dec 3: Trustmark from Ron
                    Coll.Add(New TotalsItem("AHL", "EZTmark", 52, "2011", Nothing))
                    Coll.Add(New TotalsItem("Aramark", "EZTmark", 677, "2011", Nothing))
                    Coll.Add(New TotalsItem("Martinrea", "EZTmark", 28392, "2011", Nothing))
                    Coll.Add(New TotalsItem("Morgans", "EZTmark", 10036, "2011", Nothing))
                    Coll.Add(New TotalsItem("RHSheppard", "EZTmark", 208, "2011", Nothing))
                    Coll.Add(New TotalsItem("RiteAid", "EZTmark", 260, "2011", Nothing))
                    'Value for CCIU + SEIU combined with new info received Dec 8
                    'CCIU  $ 208,763  and SEIU   $260,040.44 
                    Coll.Add(New TotalsItem("OptionsChoices", "EZTmark", 468803, "2011", Nothing))
                    Coll.Add(New TotalsItem("SFWMD", "EZTmark", 20437, "2011", Nothing))
                    Coll.Add(New TotalsItem("Vincents", "EZTmark", 52, "2011", Nothing))
                    Coll.Add(New TotalsItem("UFCW204", "EZTmark", 52, "2011", Nothing))

                    ' ___ Dec 13: Allstate from Irv
                    Coll.Add(New TotalsItem("Weathershield", "EZAllstate", 7852, "2011", Nothing))
                    Coll.Add(New TotalsItem("CTCA", "EZAllstate", 520, "2011", Nothing))
                    Coll.Add(New TotalsItem("Genesis", "EZAllstate", 63648, "2011", Nothing))

                    ' ___ Dec 13: Transamerica from Ron
                    Coll.Add(New TotalsItem("RDS", "EZTrans", 25688.0, "2011", Nothing))
                    Coll.Add(New TotalsItem("LNT", "EZTrans", 780.0, "2011", Nothing))
                    Coll.Add(New TotalsItem("PepBoys", "EZTrans", 5148.0, "2011", Nothing))
                    'Added in Stantec Trustmark EZ Value of  $12,272.00 to the Trans EZ value
                    Coll.Add(New TotalsItem("Stantec", "EZTrans", 33696, "2011", Nothing))

                    Coll.Add(New TotalsItem("BureauVeritas", "EZTrans", 13676.0, "2011", Nothing))

                Case "2012"
                    ' ___ Dec 7: Trustmark from Ron -- iupdated 1/3/2012
                    Coll.Add(New TotalsItem("AHL", "EZTmark", 52, "2012", Nothing))
                    Coll.Add(New TotalsItem("Aramark", "EZTmark", 208, "2012", Nothing))
                    Coll.Add(New TotalsItem("IPBC", "EZTmark", 5928, "2012", Nothing))
                    Coll.Add(New TotalsItem("Martinrea", "EZTmark", 46644.0, "2012", Nothing))
                    Coll.Add(New TotalsItem("Morgans", "EZTmark", 5720, "2012", Nothing))
                    Coll.Add(New TotalsItem("PKOH", "EZTmark", 30941.32, "2012", Nothing))
                    Coll.Add(New TotalsItem("RiteAid", "EZTmark", 52, "2012", Nothing))
                    'Value for CCIU + SEIU combined with new info received Jan 3
                    'CCIU  $214,900.72    and SEIU   $296,706.20 
                    Coll.Add(New TotalsItem("OptionsChoices", "EZTmark", 511606.92, "2012", Nothing))
                    Coll.Add(New TotalsItem("SFWMD", "EZTmark", 17577.36, "2012", Nothing))
                    Coll.Add(New TotalsItem("Stantec", "EZTmark", 29068.0, "2012", Nothing))
                    Coll.Add(New TotalsItem("UFCW204", "EZTmark", 156, "2012", Nothing))


                    ' ___ Dec 7: Transamerica from Ron
                    Coll.Add(New TotalsItem("RDS", "EZTrans", 23764.0, "2012", Nothing))
                    Coll.Add(New TotalsItem("LNT", "EZTrans", 156.0, "2012", Nothing))
                    Coll.Add(New TotalsItem("PepBoys", "EZTrans", 1040.0, "2012", Nothing))
                    'Added in Stantec Trustmark EZ Value of  $25480 to the Trans EZ value 14924.00
                    Coll.Add(New TotalsItem("Stantec", "EZTrans", 40404.0, "2012", Nothing))

                    Coll.Add(New TotalsItem("BureauVeritas", "EZTrans", 15028.0, "2012", Nothing))
                    Coll.Add(New TotalsItem("C3", "EZTrans", 52.0, "2012", Nothing))

                    ' ___ Dec 10: Allstate from Ron
                    Coll.Add(New TotalsItem("Weathershield", "EZAllstate", 10868, "2012", Nothing))
                    Coll.Add(New TotalsItem("CTCA", "EZAllstate", 676, "2012", Nothing))
                    Coll.Add(New TotalsItem("Genesis", "EZAllstate", 72904, "2012", Nothing))
                    Coll.Add(New TotalsItem("DioPitt", "EZAllstate", 28444, "2012", Nothing))


            End Select

            Return Coll

        Catch ex As Exception
            cReport = New Report
            cReport.Report("Error# 1456: Rpt_BVIProduction.GetSectionTwoTotalsColl " & ex.Message, Report.ReportTypeEnum.LogError)
        End Try
    End Function


    Private Function GetRangeColl(ByVal ReportYear As String) As Collection
        Dim Coll As New Collection

        Try
            Select Case ReportYear
                Case "2010"

                    ' ___ From Ron 12/21
                    Coll.Add(New RangeItem("_2010_Allstate_GNC_VIS", 1243))
                    Coll.Add(New RangeItem("_2010_Allstate_GNC_UL", 28682))
                    Coll.Add(New RangeItem("_2010_Allstate_GNC_CI", 11161))
                    Coll.Add(New RangeItem("_2010_Allstate_GNC_ACC", 17963))
            End Select
            Return Coll

        Catch ex As Exception
            cReport = New Report
            cReport.Report("Error# 1456: Rpt_BVIProduction.GetRangeColl " & ex.Message, Report.ReportTypeEnum.LogError)
        End Try
    End Function
#End Region

#Region " Helper classes "
    Public MustInherit Class Item
        Private cClientID As String
        Private cProductID As String

        Public Sub New(ByVal ClientID As String, ByVal ProductID As String)
            cClientID = ClientID
            cProductID = ProductID
        End Sub

        Public ReadOnly Property ClientID() As String
            Get
                Return cClientID
            End Get
        End Property
        Public ReadOnly Property ProductID() As String
            Get
                Return cProductID
            End Get
        End Property
    End Class

    Public Class ClientItem
        Inherits Item
        Private cCompositeClientInd As Boolean
        Private cClientStartDate As Object
        Private cClientEndDate As Object
        Private cReport As Report

        Public Sub New(ByVal ClientID As String, ByVal ProductID As String, ByVal dt As DataTable)

            MyBase.New(ClientID, ProductID)

            Dim i As Integer
            Try
                If InStr(ClientID, "|") > 0 Then
                    cCompositeClientInd = True
                End If

                For i = 0 To dt.Rows.Count - 1
                    If dt.Rows(i)(0) = ClientID Then
                        cClientStartDate = dt.Rows(i)("StartDate")
                        cClientEndDate = dt.Rows(i)("EndDate")
                        Exit For
                    End If
                Next

            Catch ex As Exception
                cReport = New Report
                cReport.Report("Error# 1457: ClientItem.New " & ex.Message, Report.ReportTypeEnum.LogError)
            End Try
        End Sub

        Public ReadOnly Property CompositeClientInd() As Boolean
            Get
                Return cCompositeClientInd
            End Get
        End Property
        Public ReadOnly Property ClientStartDate() As Date
            Get
                Return cClientStartDate
            End Get
        End Property
        Public ReadOnly Property ClientEndDate() As Object
            Get
                Return cClientEndDate
            End Get
        End Property
    End Class

    Public Class PartialYearItem
        Inherits Item
        Private cRangeName As String
        Private cValue As Decimal
        Private cExternalValueYear As String

        Public Sub New(ByVal ClientID As String, ByVal ProductID As String, ByVal RangeName As String, ByVal Value As Decimal, ByVal ExternalValueYear As String)
            MyBase.New(ClientID, ProductID)
            cRangeName = RangeName
            cValue = Value
            cExternalValueYear = ExternalValueYear
        End Sub

        Public Sub New(ByVal ClientID As String, ByVal ProductID As String, ByVal RangeName As String)
            MyBase.New(ClientID, ProductID)
            cRangeName = RangeName
        End Sub

        Public ReadOnly Property RangeName() As String
            Get
                Return cRangeName
            End Get
        End Property
        Public ReadOnly Property Value() As Decimal
            Get
                Return cValue
            End Get
        End Property
        Public ReadOnly Property ExternalValueYear() As String
            Get
                Return cExternalValueYear
            End Get
        End Property
    End Class

    Public Class TotalsItem
        Inherits PartialYearItem

        Private cColumnNameOverride As String

        Public Sub New(ByVal ClientID As String, ByVal ProductID As String, ByVal Value As Decimal, ByVal ExternalValueYear As String, ByVal ColumnNameOverride As String)
            MyBase.New(ClientID, ProductID, "_" & ExternalValueYear & "_" & ProductID, Value, ExternalValueYear)
            cColumnNameOverride = ColumnNameOverride
            'MyBase.New(ClientID, ProductID, IIf(RangeNameOverride = Nothing, "_" & ExternalValueYear & "_" & ProductID, RangeNameOverride), Value, ExternalValueYear)
        End Sub

        Public ReadOnly Property ColumnNameOverride() As String
            Get
                Return cColumnNameOverride
            End Get
        End Property
    End Class

    Public Class RangeItem
        Private cRangeName As String
        Private cValue As Integer

        Public Sub New(ByVal RangeName As String, ByVal Value As Integer)
            cRangeName = RangeName
            cValue = Value
        End Sub

        Public ReadOnly Property RangeName() As String
            Get
                Return cRangeName
            End Get
        End Property

        Public ReadOnly Property Value() As Integer
            Get
                Return cValue
            End Get
        End Property
    End Class
#End Region
End Class

