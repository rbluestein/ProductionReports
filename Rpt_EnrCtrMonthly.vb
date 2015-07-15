Public Class Rpt_EnrCtrMonthly
    Public Event NotifyForm(ByRef NotifyFormArgs As NotifyFormArgs)
    Dim e As New NotifyFormArgs(NotifyFormArgs.SourceEnum.Rpt_EnrCtrMonthly)
    Private cEnviro As Enviro
    Private cCommon As New Common
    Private cReport As Report
    Private cReportConfig As New EnrollCenterMonthlyReportConfig
    Private cExcel_Generic As New Excel_Generic(cReportConfig)
    Private cOutputFullPath As String

    Public Sub New()
        cEnviro = gEnviro
    End Sub

    Public ReadOnly Property OutputFullPath() As String
        Get
            Return cOutputFullPath
        End Get
    End Property

    Public Function Init() As Results
        Dim MyResults As New Results
        Dim ExportTableToExcelResults As New Results
        Dim dtCalendar As DataTable
        Dim dtExcel As DataTable
        Dim dtHeader As DataTable
        Dim ProcessWeekResults As Results
        Dim CurCalendarTableRow As Integer
        Dim CurWeek As Integer
        Dim StartingDayOfWeek As Integer
        Dim RowOffset As Integer
        Dim ActiveExcel As Microsoft.Office.Interop.Excel.Application
        Dim ExcelAddress As New Excel_Generic.ExcelAddress
        Dim ReportDate As Date
        Dim LastReportDate As DateTime
        Dim MonthNum As Integer

        Try

            ' ___ Output
            e.Message = "Rpt_EnrCtrMonthly start"
            RaiseEvent NotifyForm(e)

            ' ___ ReportDate
            ReportDate = cReportConfig.ReportDate

            dtExcel = GetExcelDT()

            dtCalendar = GetCalendarDT(ReportDate)

            ' ___ Report header
            dtHeader = GetHeader(dtCalendar)
            ExcelAddress.RangeName = "ReportHeader"
            ExportTableToExcelResults = cExcel_Generic.ExportTableToExcel(dtHeader, ExcelAddress)
            ActiveExcel = ExportTableToExcelResults.Value

            ' ___ Daily totals
            CurCalendarTableRow = 0
            Do

                CurWeek = dtCalendar.Rows(CurCalendarTableRow)("WeekNum")

                ' ___ Output
                e.Message = "Processing week #: " & CurWeek.ToString
                RaiseEvent NotifyForm(e)

                StartingDayOfWeek = dtCalendar.Rows(CurCalendarTableRow)("DayOfWeek")
                ProcessWeekResults = ProcessWeek(CurCalendarTableRow, dtCalendar, dtExcel)
                If Not ProcessWeekResults.Success Then
                    MyResults.Success = False
                    MyResults.Message = ProcessWeekResults.Message
                    Return MyResults
                End If

                e.Message = "Processing Week" & CType(CurWeek, System.String)
                RaiseEvent NotifyForm(e)

                RowOffset = (StartingDayOfWeek - 1) * 2

                ' ___ Write the daily records for the week to excel
                '    ExportTableToExcelResults = cExcel_Generic.ExportTableToExcel(ProcessWeekResults.Value, 1, RangeName, RowOffset, ActiveExcel)
                ExcelAddress.RangeName = "Week" & CurWeek.ToString
                'ExportTableToExcelResults = cExcel_Generic.ExportTableToExcel(ProcessWeekResults.Value, ExcelAddress, ActiveExcel)
                'ExcelAddress.RangeName = RangeName
                ExcelAddress.RowOffset = RowOffset
                ExportTableToExcelResults = cExcel_Generic.ExportTableToExcel(ProcessWeekResults.Value, ExcelAddress, ActiveExcel)

                ActiveExcel = ExportTableToExcelResults.Value
                dtExcel.Rows.Clear()

            Loop Until CurCalendarTableRow >= dtCalendar.Rows.Count


            ' ___ Output
            e.Message = "Starting monthly totals"
            RaiseEvent NotifyForm(e)

            ' ___ Monthly totals
            LastReportDate = ReportDate.AddDays(-1)
            For MonthNum = 1 To LastReportDate.Month
                ProcessCallCenter("HBG", LastReportDate, MonthNum, ActiveExcel)
                ProcessCallCenter("OKC", LastReportDate, MonthNum, ActiveExcel)
            Next


            cExcel_Generic.Finish(ActiveExcel)

            ' ___ OutputFullPath
            cOutputFullPath = cExcel_Generic.OutputFullPath

            MyResults.Success = True
            Return MyResults

        Catch ex As Exception
            cReport = New Report
            cReport.Report("Error# 1450: Rpt_EnrollCtrMonthly.Init " & ex.Message, Report.ReportTypeEnum.LogError)
        End Try
    End Function

    Private Function GetHeader(ByRef dtCalendar As DataTable) As DataTable
        Dim dt As New DataTable
        Dim dr As DataRow
        Dim Text As String
        Dim WorkingDate As DateTime

        dt.Columns.Add(New DataColumn)


        WorkingDate = dtCalendar.Rows(0)("Date")
        Text = WorkingDate.ToString("MMM d")

        If dtCalendar.Rows.Count > 1 Then
            WorkingDate = dtCalendar.Rows(dtCalendar.Rows.Count - 1)("Date")
            Text &= " To " & WorkingDate.ToString("MMM d")
        End If

        Text &= " " & WorkingDate.ToString("yyyy")
        dr = dt.NewRow
        dr(0) = Text
        dt.Rows.Add(dr)

        Return dt
    End Function

    'Private Sub HandleTotals(ByRef dtExcel As DataTable, ByRef dtWeeklyTotals As DataTable, ByRef dtGrandTotal As DataTable)
    '    Dim i As Integer
    '    Dim CallCenter As String
    '    Dim dr As DataRow
    '    Dim Filter As String

    '    For i = 1 To 2

    '        If i = 1 Then
    '            CallCenter = "HBG"
    '        Else
    '            CallCenter = "OKC"
    '        End If

    '        Filter = "CallCenter='" & CallCenter & "'"

    '        dr = dtWeeklyTotals.NewRow
    '        dr("ProcessDayStr") = String.Empty
    '        dr("CallCenter") = CallCenter
    '        dr("Interviews") = dtExcel.Compute("Sum(Interviews)", Filter)
    '        dr("Enrolled") = dtExcel.Compute("Sum(Enrolled)", Filter)

    '        dr("EnrollMD") = dtExcel.Compute("Sum(EnrollMD)", Filter)
    '        dr("AdmMD") = dtExcel.Compute("Sum(AdmMD)", Filter)
    '        dr("TotalMD") = dtExcel.Compute("Sum(TotalMD)", Filter)
    '        dr("Premium") = dtExcel.Compute("Sum(Premium)", Filter)
    '        dr("InterviewMD") = dtExcel.Compute("Sum(InterviewMD)", Filter)
    '        dr("PremiumMD") = dtExcel.Compute("Sum(PremiumMD)", Filter)

    '        dr("PremiumInterview") = dtExcel.Compute("Sum(PremiumInterview)", Filter)
    '        dr("PremiumEnrolled") = dtExcel.Compute("Sum(PremiumEnrolled)", Filter)
    '        dr("Ratio") = dtExcel.Compute("Sum(Ratio)", Filter)
    '    Next
    'End Sub

    Private Function GetCalendarDT(ByVal ReportDate As Date) As DataTable
        Dim i As Integer
        Dim FirstOfMonthDate As Date
        Dim RptMonth As Integer
        Dim RptYear As Integer
        Dim WeekNum As Integer
        Dim ProcessDate As DateTime
        Dim ProcessDayOfWeek As Integer
        Dim dtCalendar As New DataTable
        Dim dr As DataRow
        Dim LastReportDate As DateTime
        Dim LastReportDay As Integer

        Try

            ' ___ Output
            e.Message = "Building calendar"
            RaiseEvent NotifyForm(e)

            ' ___ Bulid dtCalendar
            dtCalendar.Columns.Add(New DataColumn("Date", GetType(System.DateTime)))
            dtCalendar.Columns.Add(New DataColumn("DayOfWeek", GetType(System.Int16)))
            dtCalendar.Columns.Add(New DataColumn("WeekNum", GetType(System.Int16)))

            ' ___ The last report day is the day prior to the report date
            LastReportDate = ReportDate.AddDays(-1)
            LastReportDay = LastReportDate.Day
            RptMonth = LastReportDate.Month
            RptYear = LastReportDate.Year


            FirstOfMonthDate = CType("#" & RptMonth.ToString & "/1/" & RptYear.ToString & "#", DateTime)
            '       DaysInMonth = Date.DaysInMonth(RptYear, RptMonth)
            WeekNum = 1


            ProcessDate = FirstOfMonthDate
            For i = 1 To LastReportDay

                ' ___ Monday through Sunday: 1 - 7
                ProcessDayOfWeek = CType(ProcessDate.DayOfWeek, Integer)
                If ProcessDayOfWeek = 0 Then
                    ProcessDayOfWeek = 7
                End If

                dr = dtCalendar.NewRow
                dr(0) = ProcessDate
                dr(1) = ProcessDayOfWeek
                dr(2) = WeekNum
                dtCalendar.Rows.Add(dr)

                If ProcessDayOfWeek = 7 Then
                    WeekNum += 1
                End If
                ProcessDate = ProcessDate.AddDays(1)

            Next

            Return dtCalendar

        Catch ex As Exception
            cReport = New Report
            cReport.Report("Error# 1461: Rpt_EnrollCenterMonthly.GetCalendarDT " & ex.Message, Report.ReportTypeEnum.LogError)
        End Try
    End Function

    Private Function GetExcelDT() As DataTable
        Dim dtExcel As New DataTable

        dtExcel.Columns.Add(New DataColumn("ProcessDayStr", GetType(System.String)))
        dtExcel.Columns.Add(New DataColumn("CallCenter", GetType(System.String)))
        dtExcel.Columns.Add(New DataColumn("Interviews", GetType(System.Int64)))
        dtExcel.Columns.Add(New DataColumn("Enrolled", GetType(System.Int64)))
        dtExcel.Columns.Add(New DataColumn("EnrollMD", GetType(System.Decimal)))
        dtExcel.Columns.Add(New DataColumn("AdmMD", GetType(System.Decimal)))
        dtExcel.Columns.Add(New DataColumn("TotalMD", GetType(System.Decimal)))
        dtExcel.Columns.Add(New DataColumn("Premium", GetType(System.Decimal)))
        dtExcel.Columns.Add(New DataColumn("InterviewMD", GetType(System.Decimal)))
        dtExcel.Columns.Add(New DataColumn("PremiumMD", GetType(System.Decimal)))
        dtExcel.Columns.Add(New DataColumn("PremiumInterview", GetType(System.Decimal)))
        dtExcel.Columns.Add(New DataColumn("PremiumEnrolled", GetType(System.Decimal)))
        dtExcel.Columns.Add(New DataColumn("Ratio", GetType(System.Decimal)))
        Return dtExcel
    End Function

    Private Function ProcessWeek(ByRef CurCalendarTableRow As Integer, ByRef dtCalendar As DataTable, ByRef dtExcel As DataTable) As Results
        Dim CurWeek As Integer
        Dim DayOfWeek As Integer
        Dim MyResults As New Results
        Dim ProcessDayResults As Results
        Dim ProcessDate As DateTime

        Try

            CurWeek = dtCalendar.Rows(CurCalendarTableRow)("WeekNum")
            DayOfWeek = dtCalendar.Rows(CurCalendarTableRow)("DayOfWeek")
            Do

                ProcessDate = dtCalendar.Rows(CurCalendarTableRow)("Date")
                ProcessDayResults = ProcessDay(ProcessDate, dtCalendar.Rows(CurCalendarTableRow), dtExcel)
                If Not ProcessDayResults.Success Then
                    MyResults.Success = False
                    MyResults.Message = ProcessDayResults.Message
                    Return MyResults
                End If

                CurCalendarTableRow += 1

                If CurCalendarTableRow >= dtCalendar.Rows.Count Then
                    Exit Do
                End If

            Loop Until dtCalendar.Rows(CurCalendarTableRow)("WeekNum") > CurWeek

            MyResults.Success = True
            MyResults.Value = dtExcel
            Return MyResults

        Catch ex As Exception
            cReport = New Report
            cReport.Report("Error# 1451: Rpt_EnrollCenterMonthly.ProcessWeek " & ex.Message, Report.ReportTypeEnum.LogError)
        End Try
    End Function

    Private Function ProcessDay(ByVal ProcessDate As DateTime, ByRef drCalendar As DataRow, ByRef dtExcel As DataTable) As Results
        Dim i As Integer
        Dim MyResults As New Results
        Dim dt As DataTable
        Dim dr As DataRow
        Dim CallCenter As String
        Dim Sql As New System.Text.StringBuilder
        Dim ProcessDateStr As String
        Dim dtPremium As DataTable
        Dim AnnualPremiumInd As Boolean
        Dim AnnualPremium As String
        Dim DisplayResultsInd As Boolean
        Dim HBGResults() As Object
        Dim OKCResults() As Object
        Dim Querypack As QueryPack

        Try

            For i = 1 To 2
                If i = 1 Then
                    ProcessDateStr = ProcessDate.ToString("MM/dd/yyyy")
                    CallCenter = "HBG"
                Else
                    ProcessDateStr = String.Empty
                    CallCenter = "OKC"
                End If

                AnnualPremium = "0"
                Sql.Length = 0
                Sql.Append("SELECT AnnualPremium = SUM(Cast(FieldData as decimal(10,2))) ")
                Sql.Append("FROM Rpt_ProductHistory rph ")
                Sql.Append("LEFT JOIN UserManagement..Users u on rph.EnrollerID = u.UserID ")
                Sql.Append("WHERE dbo.ufn_IsDateEqual(rph.CallDate, '" & ProcessDate & "') = 1 AND ")
                Sql.Append("u.LocationID = '" & CallCenter & "' AND rph.FieldName = 'AnnualPremium'")

                Querypack = cCommon.GetDTWithQuerypack(Sql.ToString)
                If Not Querypack.Success Then
                    MyResults.Success = False
                    MyResults.Message = Querypack.TechErrMsg
                    Return MyResults
                End If

                dtPremium = Querypack.dt

                If (Not IsDBNull(dtPremium.Rows(0)(0))) AndAlso (dtPremium.Rows(0)(0) <> 0) Then
                    AnnualPremiumInd = True
                    AnnualPremium = CType(dtPremium.Rows(0)(0), System.String)
                End If

                Sql.Length = 0
                Sql.Append("SELECT ")
                Sql.Append("ProcessDate = '" & ProcessDateStr & "', ")
                Sql.Append("CallCenter = '" & CallCenter & "', ")
                Sql.Append("Interviews = IsNull(SUM(Interviewed), 0), ")
                Sql.Append("Enrolled =  IsNull(SUM(Enrolled), 0), ")
                Sql.Append("EnrollMD = (  IsNull(SUM(EnrollHours), 0) / 8), ")
                Sql.Append("AdmMD = ((   IsNull(SUM(AdminHours), 0)  + IsNull(SUM(TrainHours), 0) + IsNull(SUM(CoachHours), 0)   ) / 8), ")
                Sql.Append("TotalMD = (( IsNull(SUM(EnrollHours), 0) + IsNull(SUM(AdminHours), 0)  + IsNull(SUM(TrainHours), 0) + IsNull(SUM(CoachHours), 0)  ) / 8), ")

                If AnnualPremiumInd Then
                    Sql.Append("Premium = " & AnnualPremium & ", ")
                Else
                    Sql.Append("Premium = 0, ")
                End If

                Sql.Append("InterviewedMD = case ")
                Sql.Append("when SUM(IsNull(EnrollHours, 0)) = 0 then 0 ")
                Sql.Append("else  IsNull(SUM(Interviewed), 0) / (SUM(EnrollHours)  / 8) ")
                Sql.Append("end, ")

                If AnnualPremiumInd Then
                    Sql.Append("PremiumMD =  case ")
                    Sql.Append("when IsNull(SUM(EnrollHours), 0) = 0 then 0 ")
                    Sql.Append("else Cast(" & AnnualPremium & " / (SUM(EnrollHours)  / 8) as decimal(10,2)) ")
                    Sql.Append("end, ")

                    Sql.Append("PremiumInterview = case ")
                    Sql.Append("when IsNull(SUM(Interviewed), 0) = 0 then 0 ")
                    Sql.Append("else Cast(" & AnnualPremium & " / SUM(Interviewed) as decimal(10,2)) ")
                    Sql.Append("end, ")

                    Sql.Append("PremiumEnrolled = case ")
                    Sql.Append("when IsNull(SUM(Enrolled), 0) = 0 then 0 ")
                    Sql.Append("else Cast(" & AnnualPremium & " / SUM(Enrolled) as decimal(10,2)) ")
                    Sql.Append("end, ")

                Else
                    Sql.Append("PremiumMD =  0, ")
                    Sql.Append("PremiumInterview = 0, ")
                    Sql.Append("PremiumEnrolled = 0, ")
                End If

                Sql.Append("Ratio = case ")
                Sql.Append("when IsNull(SUM(Interviewed), 0) = 0 then 0 ")
                Sql.Append("else (IsNull(SUM(Enrolled), 0) * 1.0)   / (SUM(Interviewed) * 1.0) ")
                Sql.Append("end ")

                Sql.Append("FROM Rpt_CallHistory rch ")
                Sql.Append("LEFT JOIN UserManagement..Users u on rch.EnrollerID = u.UserID ")
                Sql.Append("WHERE dbo.ufn_IsDateEqual(rch.CallDate, '" & ProcessDate & "') = 1 AND ")
                Sql.Append("u.LocationID = '" & CallCenter & "'")

                Querypack = cCommon.GetDTWithQuerypack(Sql.ToString)
                If Not Querypack.Success Then
                    MyResults.Success = False
                    MyResults.Message = Querypack.TechErrMsg
                    Return MyResults
                End If

                dt = Querypack.dt

                If i = 1 Then
                    HBGResults = dt.Rows(0).ItemArray
                Else
                    OKCResults = dt.Rows(0).ItemArray
                End If
            Next


            If drCalendar("DayOfWeek") < 6 Then
                DisplayResultsInd = True
            Else
                If HBGResults(2) <> 0 Or HBGResults(3) <> 0 Or OKCResults(2) <> 0 Or OKCResults(3) <> 0 Then
                    DisplayResultsInd = True
                End If
            End If

            If DisplayResultsInd Then
                dr = dtExcel.NewRow
                dr.ItemArray = HBGResults
                dtExcel.Rows.Add(dr)
                dr = dtExcel.NewRow
                dr.ItemArray = OKCResults
                dtExcel.Rows.Add(dr)
            End If

            MyResults.Success = True
            Return MyResults

        Catch ex As Exception
            cReport = New Report
            cReport.Report("Error# 1452: Rpt_EnrollCenterMonthly.ProcessDay " & ex.Message, Report.ReportTypeEnum.LogError)
        End Try
    End Function

    Private Function ProcessCallCenter(ByVal CallCenter As String, ByVal LastReportDate As Date, ByVal MonthNum As Integer, ByRef ActiveExcel As Microsoft.Office.Interop.Excel.Application) As Results
        Dim MyResults As New Results
        Dim Results As New Results
        'Dim SheetNum As Integer
        Dim Sql As New System.Text.StringBuilder
        Dim Querypack As QueryPack
        Dim Where As String
        Dim AnnualPremium As String
        Dim AnnualPremiumInd As Boolean
        Dim ExcelAddress As New Excel_Generic.ExcelAddress
        Dim StartDate As String
        Dim EndDate As String
        Dim RangeName As String
        Dim SpreadsheetColl As Collection
        Dim StartRow As Integer

        Try

            ' ___ Output
            e.Message = "Processing " & CallCenter
            RaiseEvent NotifyForm(e)

            '' ___ Sheet num
            'SheetNum = 2

            ' ___ Start and end date
            StartDate = MonthNum.ToString & "/1/" & LastReportDate.Year.ToString
            EndDate = MonthNum.ToString & "/" & Date.DaysInMonth(LastReportDate.Year, MonthNum).ToString & "/" & LastReportDate.Year.ToString

            ' ___ Range name
            RangeName = cCommon.GetMonthName(MonthNum)

            ' ___ Where
            Where = "u.LocationID = '" & CallCenter & "' AND dbo.ufn_IsDateBetween(CallDate, '" & StartDate & "', '" & EndDate & "') = 1 "

            ' ___ Annual Premium
            Sql.Append("SELECT AnnualPremium = IsNull(SUM(Cast(IsNull(FieldData, 0) as decimal(10,2))), 0) ")
            Sql.Append("FROM Rpt_ProductHistory rph ")
            Sql.Append("LEFT JOIN UserManagement..Users u on rph.EnrollerID = u.UserID ")
            Sql.Append("WHERE dbo.ufn_IsDateBetween(rph.CallDate, '" & StartDate & "', '" & EndDate & "') = 1 AND ")
            Sql.Append("u.LocationID = '" & CallCenter & "' AND rph.FieldName = 'AnnualPremium'")
            Querypack = cCommon.GetDTWithQuerypack(Sql.ToString)
            If Not Querypack.Success Then
                MyResults.Success = False
                MyResults.Message = Querypack.TechErrMsg
                Return MyResults
            End If

            If (Not IsDBNull(Querypack.dt.Rows(0)(0))) AndAlso (Querypack.dt.Rows(0)(0) <> 0) Then
                AnnualPremiumInd = True
                AnnualPremium = CType(Querypack.dt.Rows(0)(0), System.String)
            Else
                AnnualPremium = "0"
            End If

            Sql.Length = 0
            Sql.Append("SELECT DISTINCT CallCenter = '" & CallCenter & "', ")
            Sql.Append("Interviews = SUM(IsNull(Interviewed, 0)), ")
            Sql.Append("Enrolled = SUM(IsNull(Enrolled, 0)), ")
            Sql.Append("EnrolledMD= IsNull(SUM(EnrollHours), 0) / 8, ")
            Sql.Append("AdmMD=  (IsNull(SUM(AdminHours), 0)  + IsNull(SUM(TrainHours), 0)   +    IsNull(SUM(CoachHours), 0)) / 8, ")
            Sql.Append("TotalMD= (IsNull(SUM(EnrollHours), 0) + IsNull(SUM(AdminHours), 0)  + IsNull(SUM(TrainHours), 0) + IsNull(SUM(CoachHours), 0)  ) / 8, ")
            Sql.Append("Premium = " & AnnualPremium & ", ")

            Sql.Append("InterviewedMD = case ")
            Sql.Append("when SUM(IsNull(EnrollHours, 0)) = 0 then 0 ")
            Sql.Append("else  SUM(IsNull(Interviewed, 0)) / (SUM(EnrollHours)  / 8) ")
            Sql.Append("end, ")

            If AnnualPremiumInd Then
                Sql.Append("PremiumMD =  case ")
                Sql.Append("when IsNull(SUM(EnrollHours), 0) = 0 then 0 ")
                Sql.Append("else Cast(" & AnnualPremium & " / (SUM(EnrollHours)  / 8) as decimal(10,2)) ")
                Sql.Append("end, ")

                Sql.Append("PremiumInterview = case ")
                Sql.Append("when IsNull(SUM(Interviewed), 0) = 0 then 0 ")
                Sql.Append("else Cast(" & AnnualPremium & " / SUM(Interviewed) as decimal(10,2)) ")
                Sql.Append("end, ")

                Sql.Append("PremiumEnrolled = case ")
                Sql.Append("when IsNull(SUM(Enrolled), 0) = 0 then 0 ")
                Sql.Append("else Cast(" & AnnualPremium & " / SUM(Enrolled) as decimal(10,2)) ")
                Sql.Append("end, ")

            Else
                Sql.Append("PremiumMD =  0, ")
                Sql.Append("PremiumInterview = 0, ")
                Sql.Append("PremiumEnrolled = 0, ")
            End If

            Sql.Append("Ratio = case ")
            Sql.Append("when IsNull(SUM(Interviewed), 0) = 0 then 0 ")
            Sql.Append("else (IsNull(SUM(Enrolled), 0) * 1.0)   / (SUM(Interviewed) * 1.0) ")
            Sql.Append("end ")

            Sql.Append("FROM Rpt_CallHistory rch ")
            Sql.Append("INNER JOIN UserManagement..Users u ON rch.EnrollerID = u.UserID ")
            Sql.Append("WHERE " & Where)

            Querypack = cCommon.GetDTWithQuerypack(Sql.ToString)
            If Not Querypack.Success Then
                MyResults.Success = False
                MyResults.Message = Querypack.TechErrMsg
                Return MyResults
            End If

            ' ___ Prepare ExcelAddress
            SpreadsheetColl = cExcel_Generic.GetRangeData(RangeName)
            StartRow = SpreadsheetColl("xlAddr1RowNum")
            If CallCenter = "OKC" Then
                ExcelAddress.RowOffset = 1
            End If
            ExcelAddress.RangeName = RangeName
            ExcelAddress.ColumnLtr = SpreadsheetColl("xlAddr1ColumnName")
            e.Message = "Processing monthly " & CallCenter
            RaiseEvent NotifyForm(e)

            Results = cExcel_Generic.ExportTableToExcel(Querypack.dt, ExcelAddress, ActiveExcel)
            ActiveExcel = Results.Value

        Catch ex As Exception
            cReport = New Report
            cReport.Report("Error# 1481: Rpt_EnrollerProductivity.ProcessCallCenter " & ex.Message, Report.ReportTypeEnum.LogError)
        End Try
    End Function

    'Private Function ProcessCallCenter(ByVal CallCenter As String, ByVal LastReportDate As Date, ByVal MonthNum As Integer, ByRef ActiveExcel As Microsoft.Office.Interop.Excel.Application) As Results
    '    Dim MyResults As New Results
    '    Dim Results As New Results
    '    Dim SheetNum As Integer
    '    Dim Sql As New System.Text.StringBuilder
    '    Dim Querypack As QueryPack
    '    Dim Where As String
    '    Dim AnnualPremium As String
    '    Dim ExcelAddress As New Excel_Generic.ExcelAddress
    '    Dim StartDate As String
    '    Dim EndDate As String
    '    Dim RangeName As String
    '    Dim SpreadsheetColl As Collection
    '    Dim StartRow As Integer

    '    Try

    '        MessageWrite(FrmMessage.txtRpt_EnrCtrMonthly, "Processing " & CallCenter)


    '        ' ___ Output
    '        System.Diagnostics.Debug.WriteLine("MonthNum: " & MonthNum.ToString & " CallCenter: " & CallCenter)

    '        ' ___ Sheet num
    '        SheetNum = 2

    '        ' ___ Start and end date
    '        StartDate = MonthNum.ToString & "/1/" & LastReportDate.Year.ToString

    '        '''FirstOfMonthDate = CType("#" & RptMonth.ToString & "/1/" & RptYear.ToString & "#", DateTime)
    '        ''''       DaysInMonth = Date.DaysInMonth(RptYear, RptMonth)
    '        EndDate = MonthNum.ToString & "/" & Date.DaysInMonth(LastReportDate.Year, MonthNum).ToString & "/" & LastReportDate.Year.ToString

    '        ' ___ Range name
    '        RangeName = cCommon.GetMonthName(MonthNum)

    '        ' ___ Where
    '        'Where = "EnrollerID = u.UserID AND dbo.ufn_IsDateBetween(CallDate, '" & StartDate & "', '" & EndDate & "') = 1 "
    '        Where = "u.LocationID = '" & CallCenter & "' AND dbo.ufn_IsDateBetween(CallDate, '" & StartDate & "', '" & EndDate & "') = 1 "

    '        ' ___ Premium
    '        ' Premium = "(Select IsNull(SUM(Cast(IsNull(FieldData, 0) as decimal(10,2))), 0) FROM Rpt_ProductHistory WHERE " & Where & " AND FieldName = 'AnnualPremium') "

    '        Sql.Length = 0
    '        ' Sql.Append("SELECT AnnualPremium = SUM(Cast(FieldData as decimal(10,2))) ")
    '        Sql.Append("SELECT AnnualPremium = IsNull(SUM(Cast(IsNull(FieldData, 0) as decimal(10,2))), 0) ")
    '        Sql.Append("FROM Rpt_ProductHistory rph ")
    '        Sql.Append("LEFT JOIN UserManagement..Users u on rph.EnrollerID = u.UserID ")
    '        Sql.Append("WHERE dbo.ufn_IsDateBetweenl(rph.CallDate, '" & StartDate & "', '" & EndDate & "') = 1 AND ")
    '        Sql.Append("u.LocationID = '" & CallCenter & "' AND rph.FieldName = 'AnnualPremium'")
    '        Querypack = cCommon.GetDTWithQuerypack(Sql.ToString)
    '        If Not Querypack.Success Then
    '            MyResults.Success = False
    '            MyResults.Message = Querypack.TechErrMsg
    '            Return MyResults
    '        End If
    '        AnnualPremium = CType(Querypack.dt.rows(0)(0), System.String)




    '        Sql.Length = 0
    '        Sql.Append("SELECT DISTINCT CallCenter = '" & CallCenter & "', ")
    '        Sql.Append("Interviews = (SELECT IsNull(Sum(Interviewed), 0) FROM Rpt_CallHistory WHERE " & Where & "), ")
    '        Sql.Append("Enrolled = (SELECT IsNull(Sum(Enrolled), 0) FROM Rpt_CallHistory WHERE " & Where & "), ")
    '        Sql.Append("EnrolledMD = (SELECT  (IsNull(SUM(EnrollHours), 0) / 8)   FROM Rpt_CallHistory WHERE " & Where & "), ")
    '        Sql.Append("AdmMD = (SELECT   (IsNull(SUM(AdminHours), 0)  + IsNull(SUM(TrainHours), 0)   +    IsNull(SUM(CoachHours), 0)) / 8 FROM Rpt_CallHistory WHERE " & Where & "), ")
    '        Sql.Append("TotalMD = (SELECT ((IsNull(SUM(EnrollHours), 0) + IsNull(SUM(AdminHours), 0)  + IsNull(SUM(TrainHours), 0) + IsNull(SUM(CoachHours), 0)  ) / 8)  FROM Rpt_CallHistory WHERE " & Where & "), ")
    '        Sql.Append("Premium = " & AnnualPremium & ", ")






    '        Sql.Append("InterviewsMD = case ")
    '        Sql.Append("when (SELECT IsNull(Sum(EnrollHours), 0) FROM Rpt_CallHistory WHERE " & Where & ") = 0 then 0 ")
    '        Sql.Append("else ")
    '        Sql.Append("(")
    '        Sql.Append("(SELECT IsNull(Sum(Interviewed), 0) FROM Rpt_CallHistory WHERE " & Where & ") ")
    '        Sql.Append("/ ")
    '        Sql.Append("((SELECT Sum(EnrollHours) FROM Rpt_CallHistory WHERE " & Where & ") / 8) ")
    '        Sql.Append(") ")
    '        Sql.Append("end, ")

    '        Sql.Append("PremiumMD = case ")
    '        Sql.Append("when (SELECT IsNull(Sum(EnrollHours), 0) FROM Rpt_CallHistory WHERE " & Where & ") = 0 then 0 ")
    '        Sql.Append("else ")
    '        Sql.Append("(" & Premium)
    '        Sql.Append("/ ")
    '        Sql.Append("((SELECT Sum(EnrollHours) FROM Rpt_CallHistory WHERE " & Where & ") / 8) ")
    '        Sql.Append(") ")
    '        Sql.Append("end, ")

    '        Sql.Append("PremiumInterviewed =  case ")
    '        Sql.Append("when (SELECT IsNull(Sum(Interviewed), 0) FROM Rpt_CallHistory WHERE " & Where & ") = 0 then 0 ")
    '        Sql.Append("else ")
    '        Sql.Append("( ")
    '        Sql.Append(Premium)
    '        Sql.Append("/ ")
    '        Sql.Append("(SELECT Sum(Interviewed) FROM Rpt_CallHistory WHERE " & Where & ") ")
    '        Sql.Append(") ")
    '        Sql.Append("end, ")

    '        Sql.Append("PremiumEnrolled = case ")
    '        Sql.Append("when (SELECT IsNull(Sum(Enrolled), 0) FROM Rpt_CallHistory WHERE " & Where & ") = 0 then 0 ")
    '        Sql.Append("else ")
    '        Sql.Append("( ")
    '        Sql.Append(Premium)
    '        Sql.Append("/ ")
    '        Sql.Append("(SELECT Sum(Enrolled) FROM Rpt_CallHistory WHERE " & Where & ") ")
    '        Sql.Append(") ")
    '        Sql.Append("end, ")

    '        Sql.Append("Ratio = case ")
    '        Sql.Append("when (SELECT IsNull(Sum(Interviewed), 0) FROM Rpt_CallHistory WHERE " & Where & ") = 0 then 0 ")
    '        Sql.Append("else ")
    '        Sql.Append("( ")
    '        Sql.Append("(SELECT IsNull(Sum(Enrolled), 0) * 1.0 FROM Rpt_CallHistory WHERE " & Where & ") ")
    '        Sql.Append("/ ")
    '        Sql.Append("(SELECT IsNull(Sum(Interviewed), 0) * 1.0 FROM Rpt_CallHistory WHERE " & Where & ") ")
    '        Sql.Append(") ")
    '        Sql.Append("End ")

    '        Sql.Append("FROM Rpt_CallHistory rch ")
    '        Sql.Append("INNER JOIN UserManagement..Users u ON rch.EnrollerID = u.UserID ")
    '        Sql.Append("WHERE " & Where)

    '        Querypack = cCommon.GetDTWithQuerypack(Sql.ToString)
    '        If Not Querypack.Success Then
    '            MyResults.Success = False
    '            MyResults.Message = Querypack.TechErrMsg
    '            Return MyResults
    '        End If

    '        ' ___ Prepare ExcelAddress
    '        SpreadsheetColl = cExcel_Generic.GetRangeData(RangeName)
    '        StartRow = SpreadsheetColl("xlAddr1RowNum")
    '        If CallCenter = "OKC" Then
    '            ExcelAddress.RowOffset = 1
    '        End If
    '        ExcelAddress.RangeName = RangeName
    '        ExcelAddress.ColumnLtr = SpreadsheetColl("xlAddr1ColumnName")
    '        MessageWrite(FrmMessage.txtRpt_EnrCtrMonthly, "Processing monthly " & CallCenter)
    '        Results = cExcel_Generic.ExportTableToExcel(Querypack.dt, ExcelAddress, ActiveExcel)
    '        ActiveExcel = Results.Value

    '    Catch ex As Exception
    '        cReport = New Report
    '        cReport.Report("Error# 1481: Rpt_EnrollerProductivity.ProcessCallCenter " & ex.Message, Report.ReportTypeEnum.LogError)
    '    End Try
    'End Function
End Class
