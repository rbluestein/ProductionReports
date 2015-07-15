Public Class Rpt_EnrProductivity
    Private cEnviro As Enviro
    Private cCommon As New Common
    Private cReport As Report
    Private cReportConfig As New EnrollerProductivityReportConfig
    Private cExcel_Generic As New Excel_Generic(cReportConfig)
    Private cActiveExcel As Microsoft.Office.Interop.Excel.Application
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
        Dim i As Integer
        Dim MyResults As New Results
        Dim ExportTableToExcelResults As New Results
        Dim DataPack As DataPack

        Try

            If CType(cReportConfig.ReportDate, System.DateTime).DayOfWeek <> DayOfWeek.Monday Then
                MyResults.Success = True
                Return MyResults
            End If

            DataPack = GetDataPack()

            For i = 1 To 2
                If i = 1 Then
                    ProcessCallCenter("HBG", DataPack)
                Else
                    ProcessCallCenter("OKC", DataPack)
                End If
            Next

            cExcel_Generic.Finish(cActiveExcel)

            ' ___ OutputFullPath
            cOutputFullPath = cExcel_Generic.OutputFullPath

            MyResults.Success = True
            Return MyResults

        Catch ex As Exception
            cReport = New Report
            cReport.Report("Error# 1450: Rpt_EnrProductivity.Init " & ex.Message, Report.ReportTypeEnum.LogError)
        End Try
    End Function

    'Private Function ORIGProcessCallCenter(ByVal CallCenter As String, ByRef DataPack As DataPack) As Results
    '    Dim MyResults As New Results
    '    Dim Results As New Results
    '    Dim SheetNum As Integer
    '    Dim dtExcel As DataTable
    '    Dim Sql As New System.Text.StringBuilder
    '    Dim Querypack As QueryPack
    '    Dim Where As String
    '    Dim Premium As String
    '    Dim ExcelAddress As New Excel_Generic.ExcelAddress


    '    Try

    '        MessageWrite(FrmMessage.txtRpt_EnrProductivity, "Processing " & CallCenter)

    '        ' ___ Sheet num
    '        If CallCenter = "HBG" Then
    '            SheetNum = 1
    '        Else
    '            SheetNum = 2
    '        End If

    '        ' ___ Where
    '        Where = "EnrollerID = rph.EnrollerID AND dbo.ufn_IsDateBetween(CallDate, '" & DataPack.FirstReportDate.ToString("M/d/yyyy") & "', '" & DataPack.LastReportDate.ToString("M/d/yyyy") & "') = 1 "

    '        ' ___ Premium
    '        Premium = "(Select IsNull(SUM(Cast(IsNull(FieldData, 0) as decimal(10,2))), 0) FROM Rpt_ProductHistory WHERE " & Where & " AND FieldName = 'AnnualPremium') "

    '        ' ___ Week number
    '        ExcelAddress.RangeName = CallCenter & "_WeekNum"
    '        Results = cExcel_Generic.ExportFieldToExcel("Week: " & CType(DataPack.WeekNum, System.String), ExcelAddress, cActiveExcel)
    '        ' Results = cExcel_Generic.ExportFieldToExcel("Week: " & CType(DataPack.WeekNum, System.String), SheetNum, CallCenter & "_WeekNum", 0, cActiveExcel

    '        cActiveExcel = Results.Value

    '        ' ___ Week end date
    '        'Results = cExcel_Generic.ExportFieldToExcel("Week: " & DataPack.LastReportDate.ToString("MM/dd/yyyy"), SheetNum, CallCenter & "_WeekEnding", 0, cActiveExcel)
    '        ExcelAddress.RangeName = CallCenter & "_WeekEnding"
    '        Results = cExcel_Generic.ExportFieldToExcel("Week: " & CType(DataPack.WeekNum, System.String), ExcelAddress, cActiveExcel)
    '        cActiveExcel = Results.Value

    '        Sql.Append("SELECT DISTINCT rph.EnrollerID, ")
    '        Sql.Append("EnrollerName = u.LastName + ', ' + u.FirstName, ")
    '        Sql.Append("Interviews = (SELECT IsNull(Sum(Interviewed), 0) FROM Rpt_CallHistory WHERE " & Where & "), ")
    '        Sql.Append("Enrolled = (SELECT IsNull(Sum(Enrolled), 0) FROM Rpt_CallHistory WHERE " & Where & "), ")
    '        Sql.Append("EnrollHours = (SELECT IsNull(Sum(EnrollHours), 0) FROM Rpt_CallHistory WHERE " & Where & "), ")
    '        Sql.Append("AdminHours = (SELECT IsNull(Sum(AdminHours), 0) +  IsNull(Sum(TrainHours), 0) +  IsNull(Sum(CoachHours), 0) FROM Rpt_CallHistory WHERE " & Where & "), ")
    '        Sql.Append("TotalHours = (SELECT IsNull(Sum(AdminHours), 0) + IsNull(Sum(EnrollHours), 0) +  IsNull(Sum(TrainHours), 0) +  IsNull(Sum(CoachHours), 0) FROM Rpt_CallHistory WHERE " & Where & "), ")

    '        Sql.Append("Premium =  ")
    '        Sql.Append("(Select IsNull(SUM(Cast(IsNull(FieldData, 0) as decimal(10,2))), 0) ")
    '        Sql.Append("FROM Rpt_ProductHistory ")
    '        Sql.Append("WHERE " & Where & " AND ")
    '        Sql.Append("FieldName = 'AnnualPremium'), ")

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
    '        Sql.Append("(")
    '        Sql.Append(Premium)
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

    '        Sql.Append("FROM Rpt_ProductHistory rph ")
    '        Sql.Append("INNER JOIN UserManagement..Users u ON rph.EnrollerID = u.UserID ")
    '        Sql.Append("WHERE " & Where & "AND u.LocationID = '" & CallCenter & "'")
    '        Sql.Append("ORDER BY EnrollerName")

    '        Querypack = cCommon.GetDTWithQuerypack(Sql.ToString)
    '        If Not Querypack.Success Then
    '            MyResults.Success = False
    '            MyResults.Message = Querypack.TechErrMsg
    '            Return MyResults
    '        End If

    '        ' ___ Remove EnrollerID column
    '        Querypack.dt.Columns.Remove("EnrollerID")
    '        'Results = cExcel_Generic.ExportTableToExcel(Querypack.dt, SheetNum, CallCenter & "_Enrollers", 0, cActiveExcel)
    '        ExcelAddress.RangeName = CallCenter & "_Enrollers"
    '        Results = cExcel_Generic.ExportTableToExcel(Querypack.dt, ExcelAddress, cActiveExcel)
    '        cActiveExcel = Results.Value

    '    Catch ex As Exception
    '        cReport = New Report
    '        cReport.Report("Error# 1481: Rpt_EnrollerProductivity.ProcessCallCenter " & ex.Message, Report.ReportTypeEnum.LogError)
    '    End Try
    'End Function

    Private Function ProcessCallCenter(ByVal CallCenter As String, ByRef DataPack As DataPack) As Results
        Dim MyResults As New Results
        Dim Results As New Results
        Dim SheetNum As Integer
        Dim Sql As New System.Text.StringBuilder
        Dim Querypack As QueryPack
        Dim Where As String
        Dim Premium As String
        Dim ExcelAddress As New Excel_Generic.ExcelAddress

        Try

            ' ___ Sheet num
            If CallCenter = "HBG" Then
                SheetNum = 1
            Else
                SheetNum = 2
            End If

            ' ___ Where
            Where = "EnrollerID = rch.EnrollerID AND dbo.ufn_IsDateBetween(CallDate, '" & DataPack.FirstReportDate.ToString("M/d/yyyy") & "', '" & DataPack.LastReportDate.ToString("M/d/yyyy") & "') = 1 "

            ' ___ Premium
            Premium = "(Select IsNull(SUM(Cast(IsNull(FieldData, 0) as decimal(10,2))), 0) FROM Rpt_ProductHistory WHERE " & Where & " AND FieldName = 'AnnualPremium') "

            ' ___ Week number
            ExcelAddress.RangeName = CallCenter & "_WeekNum"
            Results = cExcel_Generic.ExportFieldToExcel("Week: " & CType(DataPack.WeekNum, System.String), ExcelAddress, cActiveExcel)
            ' Results = cExcel_Generic.ExportFieldToExcel("Week: " & CType(DataPack.WeekNum, System.String), SheetNum, CallCenter & "_WeekNum", 0, cActiveExcel

            cActiveExcel = Results.Value

            ' ___ Week end date
            'Results = cExcel_Generic.ExportFieldToExcel("Week: " & DataPack.LastReportDate.ToString("MM/dd/yyyy"), SheetNum, CallCenter & "_WeekEnding", 0, cActiveExcel)
            ExcelAddress.RangeName = CallCenter & "_WeekEnding"
            'Results = cExcel_Generic.ExportFieldToExcel("Week: " & CType(DataPack.WeekNum, System.String), ExcelAddress, cActiveExcel)
            Results = cExcel_Generic.ExportFieldToExcel(DataPack.FirstReportDate.ToString("MMM %d yyyy") & " to " & DataPack.LastReportDate.ToString("MMM %d yyyy"), ExcelAddress, cActiveExcel)

            cActiveExcel = Results.Value

            Sql.Append("SELECT DISTINCT rch.EnrollerID, ")
            Sql.Append("EnrollerName = u.LastName + ', ' + u.FirstName, ")
            Sql.Append("Interviews = (SELECT IsNull(Sum(Interviewed), 0) FROM Rpt_CallHistory WHERE " & Where & "), ")
            Sql.Append("Enrolled = (SELECT IsNull(Sum(Enrolled), 0) FROM Rpt_CallHistory WHERE " & Where & "), ")
            Sql.Append("EnrollHours = (SELECT IsNull(Sum(EnrollHours), 0) FROM Rpt_CallHistory WHERE " & Where & "), ")
            Sql.Append("AdminHours = (SELECT IsNull(Sum(AdminHours), 0) +  IsNull(Sum(TrainHours), 0) +  IsNull(Sum(CoachHours), 0) FROM Rpt_CallHistory WHERE " & Where & "), ")
            Sql.Append("TotalHours = (SELECT IsNull(Sum(AdminHours), 0) + IsNull(Sum(EnrollHours), 0) +  IsNull(Sum(TrainHours), 0) +  IsNull(Sum(CoachHours), 0) FROM Rpt_CallHistory WHERE " & Where & "), ")

            Sql.Append("Premium =  ")
            Sql.Append("(Select IsNull(SUM(Cast(IsNull(FieldData, 0) as decimal(10,2))), 0) ")
            Sql.Append("FROM Rpt_ProductHistory ")
            Sql.Append("WHERE " & Where & " AND ")
            Sql.Append("FieldName = 'AnnualPremium'), ")

            Sql.Append("InterviewsMD = case ")
            Sql.Append("when (SELECT IsNull(Sum(EnrollHours), 0) FROM Rpt_CallHistory WHERE " & Where & ") = 0 then 0 ")
            Sql.Append("else ")
            Sql.Append("(")
            Sql.Append("(SELECT IsNull(Sum(Interviewed), 0) FROM Rpt_CallHistory WHERE " & Where & ") ")
            Sql.Append("/ ")
            Sql.Append("((SELECT Sum(EnrollHours) FROM Rpt_CallHistory WHERE " & Where & ") / 8) ")
            Sql.Append(") ")
            Sql.Append("end, ")


            Sql.Append("PremiumMD = case ")
            Sql.Append("when (SELECT IsNull(Sum(EnrollHours), 0) FROM Rpt_CallHistory WHERE " & Where & ") = 0 then 0 ")
            Sql.Append("else ")
            Sql.Append("(")
            Sql.Append(Premium)
            Sql.Append("/ ")
            Sql.Append("((SELECT Sum(EnrollHours) FROM Rpt_CallHistory WHERE " & Where & ") / 8) ")
            Sql.Append(") ")
            Sql.Append("end, ")

            Sql.Append("PremiumInterviewed =  case ")
            Sql.Append("when (SELECT IsNull(Sum(Interviewed), 0) FROM Rpt_CallHistory WHERE " & Where & ") = 0 then 0 ")
            Sql.Append("else ")
            Sql.Append("( ")
            Sql.Append(Premium)
            Sql.Append("/ ")
            Sql.Append("(SELECT Sum(Interviewed) FROM Rpt_CallHistory WHERE " & Where & ") ")
            Sql.Append(") ")
            Sql.Append("end, ")

            Sql.Append("PremiumEnrolled = case ")
            Sql.Append("when (SELECT IsNull(Sum(Enrolled), 0) FROM Rpt_CallHistory WHERE " & Where & ") = 0 then 0 ")
            Sql.Append("else ")
            Sql.Append("( ")
            Sql.Append(Premium)
            Sql.Append("/ ")
            Sql.Append("(SELECT Sum(Enrolled) FROM Rpt_CallHistory WHERE " & Where & ") ")
            Sql.Append(") ")
            Sql.Append("end, ")

            Sql.Append("Ratio = case ")
            Sql.Append("when (SELECT IsNull(Sum(Interviewed), 0) FROM Rpt_CallHistory WHERE " & Where & ") = 0 then 0 ")
            Sql.Append("else ")
            Sql.Append("( ")
            Sql.Append("(SELECT IsNull(Sum(Enrolled), 0) * 1.0 FROM Rpt_CallHistory WHERE " & Where & ") ")
            Sql.Append("/ ")
            Sql.Append("(SELECT IsNull(Sum(Interviewed), 0) * 1.0 FROM Rpt_CallHistory WHERE " & Where & ") ")
            Sql.Append(") ")
            Sql.Append("End ")

            Sql.Append("FROM Rpt_CallHistory rch ")
            Sql.Append("INNER JOIN UserManagement..Users u ON rch.EnrollerID = u.UserID ")
            Sql.Append("WHERE " & Where & "AND u.LocationID = '" & CallCenter & "'")
            Sql.Append("ORDER BY EnrollerName")

            Querypack = cCommon.GetDTWithQuerypack(Sql.ToString)
            If Not Querypack.Success Then
                MyResults.Success = False
                MyResults.Message = Querypack.TechErrMsg
                Return MyResults
            End If

            ' ___ Remove EnrollerID column
            Querypack.dt.Columns.Remove("EnrollerID")
            'Results = cExcel_Generic.ExportTableToExcel(Querypack.dt, SheetNum, CallCenter & "_Enrollers", 0, cActiveExcel)
            ExcelAddress.RangeName = CallCenter & "_Enrollers"
            Results = cExcel_Generic.ExportTableToExcel(Querypack.dt, ExcelAddress, cActiveExcel)
            cActiveExcel = Results.Value

        Catch ex As Exception
            cReport = New Report
            cReport.Report("Error# 1481: Rpt_EnrollerProductivity.ProcessCallCenter " & ex.Message, Report.ReportTypeEnum.LogError)
        End Try
    End Function


    Private Function GetDataPack() As DataPack
        Dim ReportDate As Date
        Dim DataPack As New DataPack
        Dim FirstDayNumOfYear As Integer
        Dim Working As Integer
        Dim PriorYearDays As Integer

        ' ___ The last report day is the day prior to the report date
        ReportDate = cReportConfig.ReportDate
        DataPack.FirstReportDate = ReportDate.AddDays(-7)
        DataPack.LastReportDate = ReportDate.AddDays(-1)

        ' ___ Determine week num
        FirstDayNumOfYear = CType(CType("1/1/" & ReportDate.Year.ToString, System.DateTime).DayOfWeek, Day)
        Select Case FirstDayNumOfYear
            Case 0
                PriorYearDays = 6
            Case Else
                PriorYearDays = FirstDayNumOfYear - 1
        End Select
        Working = DataPack.FirstReportDate.DayOfYear + PriorYearDays - 1
        Working = Working \ 7
        Working += 1
        DataPack.WeekNum = Working

        Return DataPack
    End Function

    Public Class DataPack
        Private cFirstReportDate As DateTime
        Private cLastReportDate As DateTime
        Private cWeekNum As Integer
        Public Property FirstReportDate() As DateTime
            Get
                Return cFirstReportDate
            End Get
            Set(ByVal Value As DateTime)
                cFirstReportDate = Value
            End Set
        End Property
        Public Property LastReportDate() As DateTime
            Get
                Return cLastReportDate
            End Get
            Set(ByVal Value As DateTime)
                cLastReportDate = Value
            End Set
        End Property
        Public Property WeekNum() As Integer
            Get
                Return cWeekNum
            End Get
            Set(ByVal Value As Integer)
                cWeekNum = Value
            End Set
        End Property
    End Class
End Class
