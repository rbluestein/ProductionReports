Public Class Rpt_SupvMaster
    Public Event NotifyForm(ByRef NotifyFormArgs As NotifyFormArgs)
    Dim e As New NotifyFormArgs(NotifyFormArgs.SourceEnum.Rpt_SupvMaster)
    Private cEnviro As Enviro
    Private cExcel As New Excel
    Private cCommon As New Common
    Private cReport As Report
    Private cReportConfig As New SupervisorMasterReportConfig
    Private cOutputFullPath As String


    Public Sub New()
        cEnviro = gEnviro
        'cFrmMessage = FrmMessage
    End Sub

    Public ReadOnly Property OutputFullPath() As String
        Get
            Return cOutputFullPath
        End Get
    End Property

    Public Function Init(ByVal OverrideReportDate As Date) As Results
        Dim i As Integer
        Dim MyResults As New Results
        Dim ExcelClusterDataResults As Results
        Dim ExcelClusterData As New ExcelClusterData
        Dim ClusterSegmentColl As Collection
        Dim SegmentConfigureColl As Collection
        Dim ExcelPack As New ExcelPack_RptSupervisor
        Dim ExportToExcelResults As New Results
        Dim ReportDate As Date

        Try

            ' ___ Report date
            'If OverrideReportDate = Nothing Then
            '    ReportDate = cReportConfig.ReportDate
            '    If ReportDate.DayOfWeek = DayOfWeek.Monday Then
            '        ReportDate = ReportDate.AddDays(-3)
            '    Else
            '        ReportDate = ReportDate.AddDays(-1)
            '    End If
            'Else
            '    ReportDate = OverrideReportDate
            'End If

            ReportDate = OverrideReportDate


            'BuildRollupTable(ExcelPack, ReportDate)

            ' ___ Get the cluster (client) and segment (product) collections
            ExcelClusterDataResults = ExcelClusterData.GetCollections(ReportNameEnum.SupervisorReport, ReportDate)
            If Not ExcelClusterDataResults.Success Then
                MyResults.Success = False
                MyResults.Message = ExcelClusterDataResults.Message
                Return MyResults
            End If
            ClusterSegmentColl = ExcelClusterDataResults.Value(1)
            SegmentConfigureColl = ExcelClusterDataResults.Value(2)

            ' ___ Build the ExcelPack. This is a collection of items, consisting of SegmentType, ClientID, CarrierID, SegmentOffset, ReportDate, dt.
            For i = 1 To ClusterSegmentColl.Count
                BuildClientSegment(ReportDate, ExcelPack, ClusterSegmentColl(i))
                BuildProductSegmentsForThisClient(ReportDate, ExcelPack, ClusterSegmentColl(i), SegmentConfigureColl)
            Next

            ' ___ Build the rollup table
            BuildRollupTable(ExcelPack, ReportDate)

            ' ___ Build the ManDay tables
            BuildManDayTables(ExcelPack, ReportDate)

            ' ___ Build the Combined table
            'BuildCombinedTable(ExcelPack, ReportDate)

            ' ___ Pass the ExcelPack, which contains the segment datatables, to the Excel method.
            ExportToExcelResults = cExcel.ExportToExcel(ExcelPack, False, cReportConfig)
            If Not ExportToExcelResults.Success Then
                MyResults.Success = False
                MyResults.Message = ExportToExcelResults.Message
                Return MyResults
            End If

            ' ___ OutputFullPath
            cOutputFullPath = cExcel.OutputFullPath

            MyResults.Success = True
            MyResults.Value = ExportToExcelResults.Value
            Return MyResults

            Return MyResults

        Catch ex As Exception
            cReport = New Report
            cReport.Report("Error# 1250: SupervisorReport.Init " & ex.Message, Report.ReportTypeEnum.LogError)
        End Try
    End Function

    'Private Sub BuildCombinedTable(ByRef ExcelPack As ExcelPack_RptSupervisor, ByVal ReportDate As Date)
    '    Dim Sql As New System.Text.StringBuilder
    '    'Sql.Append("SELECT SUM(ISNULL(edp.AdminHours, 0))  + SUM(ISNULL(edp.EnrollHours, 0))  + SUM(ISNULL(edp.TrainHours, 0)) + SUM(ISNULL(edp.CoachHours, 0)) ")
    '    Sql.Append("SELECT SUM(ISNULL(edp.EnrollHours, 0)) ")
    '    Sql.Append("FROM Rpt_CallHistory rch ")
    '    Sql.Append("INNER JOIN EnrollerDateProject edp ON edp.ProjectDate = rch.CallDate AND rch.CallDate = '" & ReportDate & "' AND edp.ClientID = 'Combined'")
    '    'Sql.Append("SELECT Sample =  '17'")
    '    ExcelPack.Coll.Add(New ExcelPack_RptSupervisor.Item(Excel.SegmentType.Client, "Combined_Hours", Nothing, 0, Nothing, Nothing, Sql.ToString, True))
    'End Sub

    Private Sub BuildRollupTable(ByRef ExcelPack As ExcelPack_RptSupervisor, ByVal ReportDate As Date)
        Dim i As Integer
        Dim Sql As New System.Text.StringBuilder
        Dim dtClient As DataTable
        'Dim Querypack As QueryPack
        Dim ClientID As String
        Dim ClientIDAdj As String

        Try

            ' ___ Get the client list
            Sql.Append("SELECT ClientID = ClusterID, ")
            Sql.Append("IsActive = case ")
            Sql.Append("when StartDate IS NOT NULL AND dbo.ufn_IsDateBetween('" & ReportDate & "', StartDate, EndDate) = 1 then 1 ")
            Sql.Append("else 0 ")
            Sql.Append("end ")
            Sql.Append(" FROM  Excel_Cluster ")
            Sql.Append("WHERE ProdRptStatusInd = 1 ")
            Sql.Append("ORDER BY RollupSeq")
            dtClient = cCommon.GetDT(Sql.ToString)


            Sql.Length = 0
            Sql.Append("SELECT Enroller = u.LastName + ', ' + u.FirstName, ")

            For i = 0 To dtClient.Rows.Count - 1
                ClientID = dtClient.Rows(i)("ClientID")
                ClientIDAdj = Replace(ClientID, "|", "")

                If dtClient.Rows(i)("IsActive") Then
                    Sql.Append(ClientIDAdj & "_Interviewed = (SELECT Interviewed FROM Rpt_CallHistory WHERE EnrollerID = u.UserID AND ClientID = '" & ClientID & "' AND  dbo.ufn_IsDateEqual('" & ReportDate & "', CallDate) = 1), ")
                    Sql.Append(ClientIDAdj & "_Enrolled =  (SELECT Enrolled FROM Rpt_CallHistory WHERE EnrollerID = u.UserID AND  ClientID = '" & ClientID & "' AND  dbo.ufn_IsDateEqual('" & ReportDate & "', CallDate) = 1), ")
                    Sql.Append(ClientIDAdj & "_TotalHours =  (SELECT TotalHours FROM Rpt_CallHistory WHERE  EnrollerID = u.UserID AND ClientID = '" & ClientID & "' AND  dbo.ufn_IsDateEqual('" & ReportDate & "', CallDate) = 1), ")

                    If ClientID.ToUpper = "OPTIONS|CHOICES" Then
                        Sql.Append("OptionsChoices_AnnualPremium = (SELECT IsNull(SUM(Cast(IsNull(FieldData, 0) as decimal(10,2))), 0) FROM Rpt_ProductHistory rph WHERE rph.EnrollerID = u.UserID AND (rph.ClientID = 'OPTIONS' OR rph.ClientID = 'CHOICES') AND rph.CallDate = '" & ReportDate & "' AND rph.FieldName = 'AnnualPremium'), ")
                    Else
                        Sql.Append(ClientIDAdj & "_AnnualPremium = (SELECT IsNull(SUM(Cast(IsNull(FieldData, 0) as decimal(10,2))), 0) FROM Rpt_ProductHistory rph WHERE rph.EnrollerID = u.UserID AND rph.ClientID = '" & ClientID & "' AND rph.CallDate = '" & ReportDate & "' AND rph.FieldName = 'AnnualPremium'), ")
                    End If

                    If i = dtClient.Rows.Count - 1 Then
                        Sql.Append(ClientIDAdj & "_Spacer = '' ")
                    Else
                        Sql.Append(ClientIDAdj & "_Spacer = '', ")
                    End If

                Else
                    Sql.Append(ClientIDAdj & "_Interviewed = 0, ")
                    Sql.Append(ClientIDAdj & "_Enrolled = 0, ")
                    Sql.Append(ClientIDAdj & "_TotalHours = 0, ")
                    Sql.Append(ClientIDAdj & "_AnnualPremium = 0, ")
                    If i = dtClient.Rows.Count - 1 Then
                        Sql.Append(ClientIDAdj & "_Spacer = '' ")
                    Else
                        Sql.Append(ClientIDAdj & "_Spacer = '', ")
                    End If
                End If

            Next

            Sql.Append("FROM UserManagement..Users u  ")
            Sql.Append("WHERE ")
            Sql.Append("(SELECT Count (*) FROM Rpt_CallHistory rchMain WHERE rchMain.EnrollerID = u.UserID AND dbo.ufn_IsDateEqual(rchMain.CallDate, '" & ReportDate & "') = 1)  > 0 ")
            Sql.Append("ORDER BY u.LastName + u.FirstName")
            ExcelPack.Coll.Add(New ExcelPack_RptSupervisor.Item(Excel.SegmentType.Client, "Master", Nothing, 0, ReportDate, cReportConfig.LocationIDStr, Sql.ToString))

        Catch ex As Exception
            cReport = New Report
            cReport.Report("Error# 1271: SupervisorReport.BuildRollupTable " & ex.Message, Report.ReportTypeEnum.LogError)
        End Try
    End Sub

    Private Sub BuildManDayTables(ByRef ExcelPack As ExcelPack_RptSupervisor, ByVal ReportDate As Date)
        Dim Sql As String

        Try

            Sql = BuildManDayTables2(ReportDate, 0)
            ExcelPack.Coll.Add(New ExcelPack_RptSupervisor.Item(Excel.SegmentType.Client, "MandayNonLA", Nothing, 0, Nothing, Nothing, Sql))

            Sql = BuildManDayTables2(ReportDate, 1)
            ExcelPack.Coll.Add(New ExcelPack_RptSupervisor.Item(Excel.SegmentType.Client, "MandayLA", Nothing, 0, Nothing, Nothing, Sql))

        Catch ex As Exception
            cReport = New Report
            cReport.Report("Error# 1251: SupervisorReport.BuildManDayTables " & ex.Message, Report.ReportTypeEnum.LogError)
        End Try
    End Sub

    'Private Function BuildManDayTables2(ByVal ReportDate As Date, ByVal Location As Integer) As String
    '    Dim Sql As New System.Text.StringBuilder

    '    Sql.Append("SELECT EnrollerName = u.LastName + ', ' + u.FirstName, Interviewed = SUM(rch.Interviewed), Enrolled = SUM(rch.Enrolled), TotalHours = SUM(rch.TotalHours), ")

    '    If Location = 0 Then
    '        Sql.Append("PremiumWritten  = (SELECT IsNull(SUM(Cast(IsNull(FieldData, 0) as decimal(10,2))), 0) FROM Rpt_ProductHistory rph WHERE rph.EnrollerID = rch.EnrollerID AND rph.ClientID <> 'Options' AND rph.ClientID <> 'Choices' AND rph.CallDate = rch.CallDate AND rph.FieldName = 'AnnualPremium'), ")
    '    Else
    '        Sql.Append("PremiumWritten  = (SELECT IsNull(SUM(Cast(IsNull(FieldData, 0) as decimal(10,2))), 0) FROM Rpt_ProductHistory rph WHERE rph.EnrollerID = rch.EnrollerID  AND (rph.ClientID = 'Options' OR rph.ClientID = 'Choices') AND rph.CallDate = rch.CallDate AND rph.FieldName = 'AnnualPremium'), ")
    '    End If

    '    Sql.Append("AdminHours = Sum(rch.AdminHours), EnrollHours = Sum(rch.EnrollHours), TrainHours = Sum(rch.TrainHours), CoachHours = SUM(rch.CoachHours) ")
    '    Sql.Append("FROM Rpt_CallHistory rch ")
    '    Sql.Append("LEFT JOIN UserManagement..Users u ON rch.EnrollerID = u.UserID ")
    '    Sql.Append("WHERE ")
    '    Sql.Append("dbo.ufn_IsDateEqual(rch.CallDate, '" & ReportDate & "') = 1 ")

    '    If Location = 0 Then
    '        Sql.Append("AND rch.ClientID <> 'Options|Choices' ")
    '    Else
    '        Sql.Append("AND rch.ClientID = 'Options|Choices' ")
    '    End If

    '    Sql.Append("GROUP BY rch.EnrollerID, u.LastName, u.FirstName, rch.CallDate ")
    '    Sql.Append("ORDER BY u.LastName, u.FirstName")
    '    Return Sql.ToString
    'End Function

    'Private Function BuildManDayTables2(ByVal ReportDate As Date, ByVal Location As Integer) As String
    '    Dim Sql As New System.Text.StringBuilder

    '    Sql.Append("SELECT EnrollerName = u.LastName + ', ' + u.FirstName, Interviewed = SUM(rch.Interviewed), Enrolled = SUM(rch.Enrolled), ")
    '    Sql.Append("TotalHours =  ISNULL(Sum(edp.AdminHours), 0) +  ISNULL(Sum(edp.EnrollHours), 0) +  ISNULL(Sum(edp.TrainHours), 0) +  ISNULL(Sum(edp.CoachHours), 0), ")
    '    Sql.Append("PremiumWritten= IsNull(SUM(Cast(IsNull(rph.FieldData, 0) as decimal(10,2))), 0), ")
    '    Sql.Append("AdminHours = ISNULL(Sum(edp.AdminHours), 0), EnrollHours = ISNULL(Sum(edp.EnrollHours), 0), TrainHours = ISNULL(Sum(edp.TrainHours), 0), CoachHours = ISNULL(SUM(edp.CoachHours), 0) ")
    '    Sql.Append("FROM Rpt_CallHistory rch ")

    '    If Location = 0 Then
    '        Sql.Append("LEFT JOIN EnrollerDateProject edp ON rch.EnrollerID = edp.EnrollerID AND rch.CallDate = edp.ProjectDate AND  edp.ClientID NOT IN ('Options', 'Choices') ")
    '    Else
    '        Sql.Append("LEFT JOIN EnrollerDateProject edp ON rch.EnrollerID = edp.EnrollerID AND rch.CallDate = edp.ProjectDate AND edp.ClientID IN ('Options', 'Choices') ")
    '    End If

    '    Sql.Append("LEFT JOIN Rpt_ProductHistory rph ON rch.EnrollerID = rph.EnrollerID AND rch.CallDate = rph.CallDate AND rph.FieldName = 'AnnualPremium' ")
    '    Sql.Append("INNER JOIN UserManagement..Users u ON rch.EnrollerID = u.UserID ")

    '    Sql.Append("WHERE ")
    '    Sql.Append("dbo.ufn_IsDateEqual(rch.CallDate, '" & ReportDate & "') = 1 ")
    '    If Location = 0 Then
    '        Sql.Append("AND rch.ClientID <> 'Options|Choices' ")
    '    Else
    '        Sql.Append("AND rch.ClientID = 'Options|Choices' ")
    '    End If

    '    Sql.Append("GROUP BY rch.EnrollerID, u.LastName, u.FirstName, rch.CallDate ")
    '    Sql.Append("ORDER BY u.LastName, u.FirstName")
    '    Return Sql.ToString
    'End Function

    Private Function BuildManDayTables2(ByVal ReportDate As Date, ByVal Location As Integer) As String
        Dim Sql As New System.Text.StringBuilder

        Sql.Append("SELECT EnrollerName = u.LastName + ', ' + u.FirstName, Interviewed = SUM(rch.Interviewed), Enrolled = SUM(rch.Enrolled), ")
        Sql.Append("TotalHours = (SELECT SUM(ISNULL(AdminHours, 0))  + SUM(ISNULL(EnrollHours, 0))  + SUM(ISNULL(TrainHours, 0)) + SUM(ISNULL(CoachHours, 0)) ")
        Sql.Append("FROM EnrollerDateProject p WHERE p.ProjectDate = rch.CallDate AND p.EnrollerID = rch.EnrollerID AND p.ClientID <> 'Combined'), ")

        'If Location = 0 Then
        '    Sql.Append("PremiumWritten  = (SELECT IsNull(SUM(Cast(IsNull(FieldData, 0) as decimal(10,2))), 0) FROM Rpt_ProductHistory rph WHERE rph.EnrollerID = rch.EnrollerID AND rph.ClientID <> 'Options' AND rph.ClientID <> 'Choices' AND rph.CallDate = rch.CallDate AND rph.FieldName = 'AnnualPremium'), ")
        'Else
        '    Sql.Append("PremiumWritten  = (SELECT IsNull(SUM(Cast(IsNull(FieldData, 0) as decimal(10,2))), 0) FROM Rpt_ProductHistory rph WHERE rph.EnrollerID = rch.EnrollerID  AND (rph.ClientID = 'Options' OR rph.ClientID = 'Choices') AND rph.CallDate = rch.CallDate AND rph.FieldName = 'AnnualPremium'), ")
        'End If

        Sql.Append("PremiumWritten  = (SELECT IsNull(SUM(Cast(IsNull(FieldData, 0) as decimal(10,2))), 0) FROM Rpt_ProductHistory rph WHERE rph.EnrollerID = rch.EnrollerID AND ")
        Sql.Append("rph.ClientID " & IIf(Location = 0, "NOT ", "") & "IN ('Options', 'Choices') AND ")
        Sql.Append("rph.CallDate = rch.CallDate AND rph.FieldName = 'AnnualPremium' AND rph.ClientID <> 'Combined'), ")


        Sql.Append("AdminHours = (SELECT SUM(ISNULL(AdminHours, 0)) FROM EnrollerDateProject p WHERE p.ProjectDate = rch.CallDate AND p.EnrollerID = rch.EnrollerID AND p.ClientID <> 'Combined'), ")
        Sql.Append("EnrollHours = (SELECT SUM(ISNULL(EnrollHours, 0)) FROM EnrollerDateProject p WHERE p.ProjectDate = rch.CallDate AND p.EnrollerID = rch.EnrollerID AND p.ClientID <> 'Combined'), ")
        Sql.Append("TrainHours = (SELECT SUM(ISNULL(TrainHours, 0)) FROM EnrollerDateProject p WHERE p.ProjectDate = rch.CallDate AND p.EnrollerID = rch.EnrollerID AND p.ClientID <> 'Combined'), ")
        Sql.Append("CoachHours = (SELECT SUM(ISNULL(CoachHours, 0)) FROM EnrollerDateProject p WHERE p.ProjectDate = rch.CallDate AND p.EnrollerID = rch.EnrollerID AND p.ClientID <> 'Combined') ")

        Sql.Append("FROM Rpt_CallHistory rch ")
        Sql.Append("LEFT JOIN UserManagement..Users u ON rch.EnrollerID = u.UserID ")
        Sql.Append("WHERE ")
        Sql.Append("dbo.ufn_IsDateEqual(rch.CallDate, '" & ReportDate & "') = 1 AND ClientID <> 'Combined' ")

        If Location = 0 Then
            Sql.Append("AND rch.ClientID <> 'Options|Choices' ")
        Else
            Sql.Append("AND rch.ClientID = 'Options|Choices' ")
        End If

        Sql.Append("GROUP BY rch.EnrollerID, u.LastName, u.FirstName, rch.CallDate ")
        Sql.Append("ORDER BY u.LastName, u.FirstName")

        Return Sql.ToString
    End Function


    Private Sub BuildClientSegment(ByVal ReportDate As Date, ByRef ExcelPack As ExcelPack_RptSupervisor, ByRef ClusterSegmentItem As ExcelClusterData.ClusterSegmentItem)
        Dim i As Integer
        Dim ClientID As String
        Dim ClientSegmentFldList As String()
        Dim Sql As New System.Text.StringBuilder

        Try

            ClientID = ClusterSegmentItem.ClientID
            ClientSegmentFldList = Split(ClusterSegmentItem.ClientSegmentFldList, "|")

            ' ___ Output
            e.Message = "Rpt_SupervisorMaster.BuildClientSegment: " & ClientID
            RaiseEvent NotifyForm(e)

            ' ___ Start the sql
            Sql.Append("SELECT ")

            ' ___ Get the sql for each column in the segment datatable
            For i = 0 To ClientSegmentFldList.GetUpperBound(0)
                If i < ClientSegmentFldList.GetUpperBound(0) Then
                    Sql.Append(ClientSegmentFldList(i) & ", ")
                Else
                    Sql.Append(ClientSegmentFldList(i) & " ")
                End If
            Next

            ' __ Complete the sql
            Sql.Append("FROM Rpt_CallHistory WHERE ClientID = '" & ClientID & "' AND dbo.ufn_IsDateEqual(CallDate, '" & ReportDate & "') = 1 ORDER BY Name")

            ' ___ Add the segment to the ExcelPack
            ExcelPack.Coll.Add(New ExcelPack_RptSupervisor.Item(Excel.SegmentType.Client, ClientID, Nothing, 0, ReportDate, cReportConfig.LocationIDStr, Sql.ToString))

        Catch ex As Exception
            'Throw New Exception("Error #553b: ExcelOut BuildWorksheetClientSegment " & ex.Message, ex)
            cReport.Report("Rpt_SupervisorMaster.BuildClientSegment  #100 " & ex.Message, Report.ReportTypeEnum.LogError)
        End Try
    End Sub

    Private Sub BuildProductSegmentsForThisClient(ByVal ReportDate As Date, ByRef ExcelPack As ExcelPack_RptSupervisor, ByRef ClusterSegmentItem As ExcelClusterData.ClusterSegmentItem, ByRef CarrierSegmentColl As Collection)
        Dim SegmentNum As Integer
        Dim FldNum As Integer
        Dim ClientID As String
        Dim SegmentList As String()
        Dim SegmentBlendList As String()
        Dim FldList As String()
        Dim FldListColl As Collection
        Dim Sql As New System.Text.StringBuilder
        Dim SegmentConfigureItem As ExcelClusterData.SegmentConfigureItem
        Dim SegmentOffset As Integer
        Dim SegmentID As String
        Dim SegmentBlend As String
        Dim FldName As String
        Dim FldNameAdj As String
        Dim Selection As String
        Dim CompositeClientInd As Boolean
        Dim SubClientID As String
        Dim SubProductID As String
        Dim dt As DataTable

        Try

            ClientID = ClusterSegmentItem.ClientID

            If ClusterSegmentItem.SegmentList = Nothing Then
                Exit Sub
            End If

            SegmentList = Split(ClusterSegmentItem.SegmentList, "|")
            SegmentBlendList = Split(ClusterSegmentItem.SegmentBlendList, "|")
            SegmentOffset = ClusterSegmentItem.FldCount + 1

            If ClusterSegmentItem.SegmentList.Length = 0 Then
                Exit Sub
            End If

            ' ___ Composite client?
            If InStr(ClientID, "|") > 0 Then
                CompositeClientInd = True
            End If

            ' // Product loop top
            For SegmentNum = 0 To SegmentList.GetUpperBound(0)

                ' ___ Prepare the helper objects for this product
                SegmentID = SegmentList(SegmentNum)

                If CompositeClientInd Then
                    SubClientID = SegmentID.Substring(0, InStr(SegmentID, "~") - 1)
                    SubProductID = cCommon.Right(SegmentID, SegmentID.Length - InStr(SegmentID, "~"))
                Else
                    SubClientID = ClientID
                    SubProductID = SegmentID
                End If

                SegmentBlend = SegmentBlendList(SegmentNum)
                SegmentConfigureItem = CarrierSegmentColl(SubProductID)
                FldListColl = New Collection

                ' ___ Columns override 3/1/2011 for C3
                dt = cCommon.GetDT("SELECT OverrideInd FROM Excel_ClusterSegment WHERE ClusterID = '" & ClientID & "' AND SegmentID = '" & SegmentID & "'")
                If dt.Rows(0)(0) Then
                    dt = cCommon.GetDT("SELECT Columns FROM Excel_SegmentConfigure_Override WHERE ClusterID = '" & ClientID & "' AND SegmentID = '" & SegmentID & "'")
                    FldList = dt.Rows(0)(0).Split("|")
                Else
                    FldList = Split(SegmentConfigureItem.FldList, "|")
                End If


                ' ___ Output
                e.Message = "Rpt_SupervisorMaster.BuildProductSegmentsForThisClient: " & SegmentID
                RaiseEvent NotifyForm(e)

                ' ___ Start the sql for this product
                Sql.Length = 0
                Sql.Append("SELECT ")

                ' ___ Cycle through the fields for this product for this client for this date
                ' // Field loop top
                For FldNum = 0 To FldList.GetUpperBound(0)
                    FldName = FldList(FldNum)

                    If InStr(FldName, "+") > 0 Then
                        FldNameAdj = "[" & FldName & "]"
                    Else
                        FldNameAdj = FldName
                    End If

                    If FldName = "AnnualPremium" Then
                        Selection = " SUM(Convert(decimal(9,2), rph.FieldData))"
                    ElseIf FldName = "WeeklyPremium" Then
                        FldName = "AnnualPremium"
                        Selection = " (SUM(Convert(decimal(9,2), rph.FieldData)) / 52)"
                    Else
                        Selection = "SUM(dbo.ufn_ToInt(rph.FieldData))"
                    End If

                    Sql.Append(FldNameAdj & " = (SELECT " & Selection)
                    Sql.Append(" FROM Rpt_ProductHistory rph ")
                    'Sql.Append("WHERE rph.FieldName = '" & FldName & "' AND rph.EnrollerID = rch.EnrollerID AND rph.ClientID = '" & ClientID & "' AND rph.ProductID = '" & SegmentID & "' AND dbo.ufn_IsDateEqual(rph.CallDate, '" & ReportDate & "') = 1) ")
                    Sql.Append("WHERE rph.FieldName = '" & FldName & "' AND rph.EnrollerID = rch.EnrollerID AND rph.ClientID = '" & SubClientID & "' AND " & GetSegmentIDAdj(SubProductID, SegmentBlend) & " AND dbo.ufn_IsDateEqual(rph.CallDate, '" & ReportDate & "') = 1) ")

                    If FldNum < FldList.GetUpperBound(0) Then
                        Sql.Append(", ")
                    Else
                        Sql.Append(" ")
                    End If
                Next
                ' // Field loop bottom

                ' __ Complete the sql
                Sql.Append("FROM Rpt_CallHistory rch WHERE rch.ClientID = '" & ClientID & "' AND dbo.ufn_IsDateEqual(rch.CallDate, '" & ReportDate & "') = 1 ORDER BY rch.Name")


                ' ___ Add the datatable to the ExcelPack
                ExcelPack.Coll.Add(New ExcelPack_RptSupervisor.Item(Excel.SegmentType.Carrier, ClientID, SegmentList(SegmentNum), SegmentOffset, Nothing, Nothing, Sql.ToString))

                ' ___ Determine the segment offset for the next segment
                SegmentOffset = SegmentOffset + FldList.GetUpperBound(0) + 2

            Next
            ' // Product loop bottom

        Catch ex As Exception
            'Throw New Exception("Error #560c: ExcelOut BuildProductSegmentsForThisClient. ClientID: " & ClientID & " SegmentID: " & SegmentID & ". FldName: " & FldName & " " & ex.Message, ex)
            cReport.Report("Rpt_SupervisorMaster.BuildProductSegmentsForThisClient#100: ClientID: " & ClientID & " SegmentID: " & SegmentID & ". FldName: " & FldName & " ex.Message " & ex.Message, Report.ReportTypeEnum.LogError)
        End Try
    End Sub

    Private Function GetSegmentIDAdj(ByVal SegmentID As String, ByVal SegmentBlend As String) As String
        Dim i As Integer
        Dim Results As New System.Text.StringBuilder
        Dim Box() As String

        If SegmentBlend.Length = 0 Then
            Return "rph.ProductID = '" & SegmentID & "' "
        Else
            Box = Split(SegmentBlend, "~")
            Results.Append("(")
            For i = 0 To Box.GetUpperBound(0)
                If i > 0 Then
                    Results.Append(" OR ")
                End If
                Results.Append("rph.ProductID = '" & Box(i) & "' ")
            Next
            Results.Append(")")
        End If

        Return Results.ToString
    End Function
End Class
