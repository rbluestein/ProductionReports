Public Class ExcelClusterData
#Region " Declarations "
    Private cCommon As Common
    Private cEnviro As Enviro
    Private cReport As Report
#End Region

#Region " Methods "
    Public Sub New()
        cEnviro = gEnviro
    End Sub

    Public Function GetCollections(ByVal ReportName As ReportNameEnum, ByVal ReportDate As Date) As Results
        Dim MyResults As New Results
        Dim Coll As New Collection
        cCommon = New Common
        Coll.Add(GetClusterSegmentColl(ReportDate))
        Coll.Add(GetSegmentConfigureColl())

        ' cSegmentBlendList = GetSegmentBlendList(ClientID)

        MyResults.Success = True
        MyResults.Value = Coll
        Return MyResults
    End Function

    Private Function GetClusterSegmentColl(ByVal ReportDate As Date) As Collection
        Dim i As Integer
        Dim dt As DataTable
        Dim ClusterSegmentColl As New Collection
        Dim Sql As New System.Text.StringBuilder

        Try

            'dt = cCommon.GetDT("SELECT DISTINCT ClusterID FROM Excel_ClusterSegment WHERE ReportID = 1 ORDER BY ClusterID")
            'dt = cCommon.GetDT("SELECT ClusterID FROM Excel_Cluster Where ProdRptStatusInd = 1 ORDER BY ClusterID")


            Sql.Append("SELECT ClusterID FROM Excel_Cluster ")
            Sql.Append("WHERE ProdRptStatusInd = 1 AND ")

            ' Sql.Append(" ClusterID = 'C3' AND ")

            Sql.Append("StartDate IS NOT NULL AND dbo.ufn_DateCompare(StartDate, '" & ReportDate & "', 1) < 1 AND ")
            Sql.Append("(EndDate IS NULL OR dbo.ufn_DateCompare(EndDate, '" & ReportDate & "', 1) > -1) ")
            Sql.Append("ORDER BY ClusterID")
            dt = cCommon.GetDT(Sql.ToString)

            'dt = cCommon.GetDT("SELECT ClusterID FROM Excel_Cluster Where ClusterID = 'C3'")


            For i = 0 To dt.Rows.Count - 1
                ClusterSegmentColl.Add(New ClusterSegmentItem(dt.Rows(i)(0)))
            Next

            'ClusterSegmentColl.Add(New ClusterSegmentItem("Armstrong"))
            'ClusterSegmentColl.Add(New ClusterSegmentItem("BureauVeritas"))
            'ClusterSegmentColl.Add(New ClusterSegmentItem("COKC"))
            'ClusterSegmentColl.Add(New ClusterSegmentItem("CTCA"))
            'ClusterSegmentColl.Add(New ClusterSegmentItem("Fulton"))
            'ClusterSegmentColl.Add(New ClusterSegmentItem("Genesis"))
            'ClusterSegmentColl.Add(New ClusterSegmentItem("HardRock"))
            'ClusterSegmentColl.Add(New ClusterSegmentItem("HT"))
            'ClusterSegmentColl.Add(New ClusterSegmentItem("Martinrea"))
            'ClusterSegmentColl.Add(New ClusterSegmentItem("Morgans"))
            'ClusterSegmentColl.Add(New ClusterSegmentItem("PKOH"))
            'ClusterSegmentColl.Add(New ClusterSegmentItem("RDS"))
            'ClusterSegmentColl.Add(New ClusterSegmentItem("Stantec"))
            'ClusterSegmentColl.Add(New ClusterSegmentItem("Superior"))
            'ClusterSegmentColl.Add(New ClusterSegmentItem("SFWMD"))
            'ClusterSegmentColl.Add(New ClusterSegmentItem("WeatherShield"))
            'ClusterSegmentColl.Add(New ClusterSegmentItem("YMCA"))

            Return ClusterSegmentColl


        Catch ex As Exception
            cReport = New Report
            cReport.Report("Error #2452: ExcelClusterData BuildClusterSegmentColl. " & ex.Message, Report.ReportTypeEnum.LogError)
        End Try
    End Function

    Private Function GetSegmentConfigureColl() As Collection
        Dim i As Integer
        Dim dt As DataTable
        Dim SegmentConfigureColl As New Collection
        Try

            dt = cCommon.GetDT("SELECT SegmentID, Columns FROM ProjectReports..Excel_SegmentConfigure WHERE ReportID = 1 AND SegmentType = 'PRODUCT'")
            For i = 0 To dt.Rows.Count - 1
                SegmentConfigureColl.Add(New SegmentConfigureItem(dt.Rows(i)("SegmentID"), dt.Rows(i)("Columns")), dt.Rows(i)("SegmentID"))
            Next

            Return SegmentConfigureColl

        Catch ex As Exception
            cReport = New Report
            cReport.Report("Error #2453: ExcelClusterData BuildSegmentConfigureColl. " & ex.Message, Report.ReportTypeEnum.LogError)
        End Try
    End Function
#End Region

#Region " ClusterSegmentItem class "
    Public Class ClusterSegmentItem
        Private cCommon As Common
        Private cClientID As String
        Private cReport As New Report
        Private cClientSegmentFldList As String
        Private cSegmentList As String
        Private cSegmentBlendList As String
        Private cFldCount As Integer


        Public Sub New(ByVal ClientID As String)
            cCommon = New Common
            Dim FldArr() As String
            cClientID = ClientID
            cClientSegmentFldList = GetClientSegmentFieldList(ClientID)
            cSegmentList = GetSegmentList(ClientID)
            cSegmentBlendList = GetSegmentBlendList(ClientID)
            FldArr = Split(cClientSegmentFldList, "|")
            cFldCount = FldArr.GetUpperBound(0) + 1
        End Sub

        Private Function GetClientSegmentFieldList(ByVal SegmentID As String) As String
            Dim dt As DataTable

            Try
                dt = cCommon.GetDT("SELECT Columns FROM ProjectReports..Excel_SegmentConfigure WHERE SegmentID = '" & ClientID & "'")
                Return dt.Rows(0)(0)

            Catch ex As Exception
                'Throw New Exception("Error #3351: ClusterSegmentItem GetFieldList. " & ex.Message, ex)
                cReport.Report("ExcelClusterData.GetClientSegmentFieldList  #100 " & ex.Message, Report.ReportTypeEnum.LogError)
            End Try
        End Function

        'Private Function GetSegmentList(ByVal ClientID As String) As String
        '    Dim i As Integer
        '    Dim dt As DataTable
        '    Dim SegmentString As String

        '    Try

        '        dt = cCommon.GetDT("SELECT SegmentID FROM ProjectReports..Excel_ClusterSegment WHERE ClusterID = '" & ClientID & "' ORDER BY SpreadsheetSeq")
        '        For i = 0 To dt.Rows.Count - 1
        '            SegmentString &= dt.Rows(i)(0) & "|"
        '        Next
        '        SegmentString = SegmentString.Substring(0, SegmentString.Length - 1)
        '        Return SegmentString

        '    Catch ex As Exception
        '        'Throw New Exception("Error #3350: ClusterSegmentItem GetSegmentList. " & ex.Message, ex)
        '        cReport.Report("ExcelClusterData.GetSegmentList  #100 " & ex.Message, Report.ReportTypeEnum.LogError)
        '    End Try
        'End Function

        'Private Function GetSegmentList(ByVal ClientID As String) As String
        '    Dim i, j As Integer
        '    Dim dt As DataTable
        '    Dim SegmentString As String
        '    Dim ThisSegment As String
        '    Dim SegmentRule As String
        '    Dim SegmentRuleInd As Boolean
        '    Dim Box1() As String
        '    Dim Box2() As String
        '    Dim DateNow As Date
        '    Dim StartDate As Object
        '    Dim EndDate As Object

        '    Try

        '        dt = cCommon.GetDT("SELECT SegmentID, SegmentRule FROM ProjectReports..Excel_ClusterSegment WHERE ClusterID = '" & ClientID & "' ORDER BY SpreadsheetSeq")
        '        For i = 0 To dt.Rows.Count - 1
        '            SegmentRuleInd = False

        '            ThisSegment = dt.Rows(i)("SegmentID")
        '            If Not IsDBNull(dt.Rows(i)("SegmentRule")) Then
        '                SegmentRuleInd = True
        '                SegmentRule = dt.Rows(i)("SegmentRule")
        '            End If

        '            If SegmentRuleInd Then
        '                DateNow = cCommon.GetServerDateTime

        '                ' ___ Create an array of segments with their start and end dates
        '                Box1 = Split(SegmentRule, "|")

        '                ' ___ Iterate through segments to determine which one is currently in effect
        '                For j = 0 To Box1.GetUpperBound(0)
        '                    Box2 = Split(Box1(j), "~")
        '                    ThisSegment = Box2(0)
        '                    StartDate = Box2(1)
        '                    EndDate = Box2(2)

        '                    If cCommon.IsDateBetween(StartDate, EndDate, DateNow) Then
        '                        Exit For
        '                    End If
        '                Next
        '            End If

        '            SegmentString &= ThisSegment & "|"

        '        Next

        '        SegmentString = SegmentString.Substring(0, SegmentString.Length - 1)
        '        Return SegmentString

        '    Catch ex As Exception
        '        Throw New Exception("Error #3350: ClusterSegmentItem GetSegmentList. " & ex.Message, ex)
        '    End Try
        'End Function

        Private Function GetSegmentList(ByVal ClientID As String) As String
            Dim i As Integer
            Dim dt As DataTable
            Dim SegmentString As String

            Try

                dt = cCommon.GetDT("SELECT SegmentID FROM ProjectReports..Excel_ClusterSegment WHERE ClusterID = '" & ClientID & "' ORDER BY SpreadsheetSeq")

                If dt.Rows.Count > 0 Then
                    For i = 0 To dt.Rows.Count - 1
                        SegmentString &= dt.Rows(i)(0) & "|"
                    Next
                    SegmentString = SegmentString.Substring(0, SegmentString.Length - 1)
                End If
                Return SegmentString

            Catch ex As Exception
                Throw New Exception("Error #3350: ClusterSegmentItem GetSegmentList. " & ex.Message, ex)
            End Try
        End Function

        Private Function GetSegmentBlendList(ByVal ClientID As String) As String
            Dim i As Integer
            Dim dt As DataTable
            Dim SegmentBlendString As String

            Try

                dt = cCommon.GetDT("SELECT SegmentBlend FROM ProjectReports..Excel_ClusterSegment WHERE ClusterID = '" & ClientID & "' ORDER BY SpreadsheetSeq")

                If dt.Rows.Count > 0 Then
                    For i = 0 To dt.Rows.Count - 1
                        If IsDBNull(dt.Rows(i)(0)) Then
                            SegmentBlendString &= "|"
                        Else
                            SegmentBlendString &= dt.Rows(i)(0) & "|"
                        End If

                    Next
                    SegmentBlendString = SegmentBlendString.Substring(0, SegmentBlendString.Length - 1)
                End If

                Return SegmentBlendString

            Catch ex As Exception
                Throw New Exception("Error #3350: ClusterSegmentItem GetSegmentList. " & ex.Message, ex)
            End Try
        End Function

        Public ReadOnly Property ClientID() As String
            Get
                Return cClientID
            End Get
        End Property
        Public ReadOnly Property FldCount() As Integer
            Get
                Return cFldCount
            End Get
        End Property
        Public ReadOnly Property ClientSegmentFldList() As String
            Get
                Return cClientSegmentFldList
            End Get
        End Property
        Public ReadOnly Property SegmentList() As String
            Get
                Return cSegmentList
            End Get
        End Property
        Public ReadOnly Property SegmentBlendList() As String
            Get
                Return cSegmentBlendList
            End Get
        End Property
    End Class
#End Region

#Region " SegmentConfigureItem class "
    Public Class SegmentConfigureItem
        Private cCarrierID As String
        Private cFldList As String
        Private cFldCount As Integer

        Public Sub New(ByVal CarrierID As String, ByVal FldList As String)
            Dim FldArr() As String
            cCarrierID = CarrierID
            cFldList = FldList
            FldArr = Split(FldList, "|")
            cFldCount = FldArr.GetUpperBound(0) + 1
        End Sub
        Public ReadOnly Property FldList() As String
            Get
                Return cFldList
            End Get
        End Property
        Public ReadOnly Property CarrierID() As String
            Get
                Return cCarrierID
            End Get
        End Property
        Public ReadOnly Property FldCount() As Integer
            Get
                Return cFldCount
            End Get
        End Property
    End Class
#End Region
End Class