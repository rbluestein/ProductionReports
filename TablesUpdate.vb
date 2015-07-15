Public Class TablesUpdate
#Region " Declarations "
    Public Event NotifyForm(ByRef NotifyFormArgs As NotifyFormArgs)
    Dim e As New NotifyFormArgs(NotifyFormArgs.SourceEnum.TablesUpdate)
    Private cEnviro As Enviro
    Private cCommon As New Common
    Private cReport As Report
    Private cTempSql2Length As Integer
    Private cLastSql As String = String.Empty
    Private cErrDetail As String = String.Empty
    Private cCallType As CallTypeEnum
    Private cUpdateDates As String
#End Region

#Region " Enums "
    Private Enum ProductTypeEnum
        Standard = 1
        Extended = 2
    End Enum

    Private Enum UpdateTypeEnum
        Undetermined = 0
        RptControlIsEmpty = 1
        SuccessfulUpdateVeryLastSession = 2
        SuccessfulUpdatePriorIteration = 3
        SuccessfulUpdateMoreThanOneSessionAgo = 4
        NeverBeenASuccessfulUpdate = 5
    End Enum

    Public Enum CallTypeEnum
        FullUpdate = 1
        UpdateDates = 2
    End Enum
#End Region

#Region " Constructor "
    Public Sub New(ByVal CallType As CallTypeEnum)
        cCallType = CallType
        cEnviro = gEnviro
    End Sub
#End Region

#Region " Init & Helpers "
    Public Function Init() As Results


        Update_Rpt_ProductHistory_ThisClient_ThisDate_ThisProduct("02/28/2012", "C3", "TRANSCCI", "06/18/2012")

        Dim i, k As Integer
        Dim MyResults As New Results
        Dim Results As Results
        Dim FirstCallDate As Date
        Dim NewAddDate As DateTime
        Dim dtCallDate As DataTable
        Dim RptControlIsPopulated As Boolean
        Dim RecID As Integer
        Dim Querypack As QueryPack
        Dim CrashCallDate As DateTime
        Dim ThisUpdateSuccessfulInd As Boolean
        Dim UpdateType As UpdateTypeEnum
        Dim FirstTimeThrough As Boolean
        Dim TriggerDate As DateTime
        Dim ErrorMessage As String
        Dim OKToProceed As Boolean
        Dim Sql As String
        Dim FailReason As String = String.Empty
        Dim ApprovalError As String
        Dim CallDate As DateTime
        Dim sb As New System.Text.StringBuilder
        Dim dtTemp As DataTable
        Dim TestIDs As New System.Text.StringBuilder

        Try

            ' ___ Output
            If cCallType = CallTypeEnum.FullUpdate Then
                e.Message = "InitializeAndUpdateReportTables"
                RaiseEvent NotifyForm(e)
            End If

            ' ___ FirstCallDate
            FirstCallDate = "8/3/2009"

            ' ___ NewAddDate
            NewAddDate = cCommon.GetServerDateTime

            For i = 1 To 10

                Try

                    If cCallType = CallTypeEnum.FullUpdate Then
                        e.Message = "Main.InitializeAndUpdateReportTables: Try #" & i.ToString
                        RaiseEvent NotifyForm(e)
                    End If

                    ResetVariables(i, UpdateType, RptControlIsPopulated, TriggerDate, CrashCallDate, OKToProceed, ErrorMessage, FirstTimeThrough)

                    ' // Determine UpdateType

                    IsRpt_ControlPopulated(i, RptControlIsPopulated, OKToProceed, ErrorMessage, FailReason)

                    If OKToProceed AndAlso i = 1 Then
                        GetUpdateTypeAndTriggerDateIfFirstTimeThrough(RptControlIsPopulated, UpdateType, TriggerDate, OKToProceed, ErrorMessage, FailReason)
                    End If

                    If OKToProceed And UpdateType = UpdateTypeEnum.Undetermined And Not FirstTimeThrough Then
                        GetUpdateTypeAndCrashCallDateIfCallDatesUpdatedLastIteration(i, NewAddDate, UpdateType, CrashCallDate, OKToProceed, ErrorMessage, FailReason)
                    End If

                    If OKToProceed And UpdateType = UpdateTypeEnum.Undetermined Then
                        WasThereEverASuccessfulUpdate(i, UpdateType, TriggerDate, OKToProceed, ErrorMessage, FailReason)
                    End If


                    If OKToProceed Then

                        Select Case UpdateType
                            Case UpdateTypeEnum.RptControlIsEmpty
                                Results = PerformCallDateDTTreatment(dtCallDate, UpdateType, FirstCallDate, FirstCallDate, Nothing)
                                If Results.Success Then
                                    If cCallType = CallTypeEnum.UpdateDates Then
                                        MyResults.Message = Results.Message
                                        Return MyResults
                                    End If
                                Else
                                    OKToProceed = False
                                    ErrorMessage = "Try #" & i.ToString & Results.Message
                                    FailReason = Results.Message
                                End If
                                If OKToProceed Then
                                    Results = CleanUpHistoryTables("DELETE [tablename] ")
                                    If Not Results.Success Then
                                        OKToProceed = False
                                        ErrorMessage = "Try #" & i.ToString & Results.Message
                                        FailReason = Results.Message
                                    End If
                                End If

                            Case UpdateTypeEnum.SuccessfulUpdateVeryLastSession
                                Results = PerformCallDateDTTreatment(dtCallDate, UpdateType, FirstCallDate, TriggerDate, Nothing)
                                If Results.Success Then
                                    If cCallType = CallTypeEnum.UpdateDates Then
                                        MyResults.Message = Results.Message
                                        Return MyResults
                                    End If
                                Else
                                    OKToProceed = False
                                    ErrorMessage = "Try #" & i.ToString & Results.Message
                                    FailReason = Results.Message
                                End If
                                ' CleanUpHistoryTables("no action")

                            Case UpdateTypeEnum.SuccessfulUpdatePriorIteration
                                Results = PerformCallDateDTTreatment(dtCallDate, UpdateType, FirstCallDate, Nothing, CrashCallDate)
                                If Results.Success Then
                                    If cCallType = CallTypeEnum.UpdateDates Then
                                        MyResults.Message = Results.Message
                                        Return MyResults
                                    End If
                                Else
                                    OKToProceed = False
                                    ErrorMessage = "Try #" & i.ToString & Results.Message
                                    FailReason = Results.Message
                                End If
                                If OKToProceed Then
                                    Results = CleanUpHistoryTables("DELETE [tablename] WHERE dbo.ufn_IsDateEqual(CallDate, '" & CrashCallDate & "') = 1")
                                    If Not Results.Success Then
                                        OKToProceed = False
                                        ErrorMessage = "Try #" & i.ToString & Results.Message
                                        FailReason = Results.Message
                                    End If
                                End If

                            Case UpdateTypeEnum.SuccessfulUpdateMoreThanOneSessionAgo
                                Results = PerformCallDateDTTreatment(dtCallDate, UpdateType, FirstCallDate, TriggerDate, Nothing)
                                If Results.Success Then
                                    If cCallType = CallTypeEnum.UpdateDates Then
                                        MyResults.Message = Results.Message
                                        Return MyResults
                                    End If
                                Else
                                    OKToProceed = False
                                    ErrorMessage = "Try #" & i.ToString & Results.Message
                                    FailReason = Results.Message
                                End If
                                If OKToProceed Then
                                    'Results = CleanUpHistoryTables("DELETE [tablename] WHERE dbo.ufn_DateCompare(CallDate, '" & TriggerDate & "') = 1")
                                    Results = CleanUpHistoryTables("DELETE [tablename] WHERE dbo.ufn_DateCompare(CallDate, '" & TriggerDate & "', 1) = 1")
                                    If Not Results.Success Then
                                        OKToProceed = False
                                        ErrorMessage = "Try #" & i.ToString & Results.Message
                                        FailReason = Results.Message
                                    End If
                                End If

                            Case UpdateTypeEnum.NeverBeenASuccessfulUpdate
                                Results = PerformCallDateDTTreatment(dtCallDate, UpdateType, FirstCallDate, FirstCallDate, Nothing)
                                If Results.Success Then
                                    If cCallType = CallTypeEnum.UpdateDates Then
                                        MyResults.Message = Results.Message
                                        Return MyResults
                                    End If
                                Else
                                    OKToProceed = False
                                    ErrorMessage = "Try #" & i.ToString & Results.Message
                                    FailReason = Results.Message
                                End If
                                If OKToProceed Then
                                    Results = CleanUpHistoryTables("DELETE [tablename] ")
                                    If Not Results.Success Then
                                        OKToProceed = False
                                        ErrorMessage = "Try #" & i.ToString & Results.Message
                                        FailReason = Results.Message
                                    End If
                                End If

                        End Select



                    End If



                    If OKToProceed Then

                        ' ************************************************************************************************************

                        ' ___ Delete test records
                        sb.Length = 0
                        sb.Append("SELECT et.CallStartTime, et.ActivityID, et.EnrollerID, ept.AppID, ept.AltProductDataID ")
                        sb.Append("FROM ProjectReports..EmpTransmittal et ")
                        sb.Append("LEFT JOIN ProjectReports..EmpProductTransmittal ept on et.ActivityID = ept.ActivityID ")

                        'sb.Append("WHERE dbo.ufn_IsDateEqual(et.CallStartTime, '" & CallDate & "') = 1 AND ")
                        sb.Append("WHERE dbo.ufn_IsDateBetween(et.CallStartTime, '" & dtCallDate.Rows(0)(0) & "', '" & dtCallDate.Rows(dtCallDate.Rows.Count - 1)(0) & "') = 1 AND ")

                        sb.Append("ProjectReports.dbo.ufn_IsTestID(et.Clientid, et.EmpID) = 1 ")
                        sb.Append("ORDER BY et.CallStartTime")

                        Querypack = cCommon.GetDTWithQuerypack(sb.ToString)
                        dtTemp = Querypack.dt

                        If dtTemp.Rows.Count > 0 Then

                            ' ___ Display in message box and stop
                            For k = 0 To dtTemp.Rows.Count - 1
                                If k = 0 Then
                                    If k = 0 Then
                                        TestIDs.Append(dtTemp.Rows(k)("ActivityID"))
                                    Else
                                        TestIDs.Append(", " & dtTemp.Rows(k)("ActivityID"))
                                    End If
                                End If
                            Next
                            'MessageBox.Show("These TestID's detected: " & TestIDs.ToString, "Test IDs", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1)
                            'Stop

                            For k = 0 To dtTemp.Rows.Count - 1
                                If Not IsDBNull(dtTemp.Rows(k)("AltProductDataID")) Then
                                    Querypack = cCommon.ExecuteNonQueryWithQuerypack("DELETE Alt_ProductData WHERE AltProductDataID = " & dtTemp.Rows(k)("AltProductDataID"))
                                End If
                                If Not IsDBNull(dtTemp.Rows(k)("AppID")) Then
                                    Querypack = cCommon.ExecuteNonQueryWithQuerypack("DELETE EmpProductTransmittal WHERE AppID = '" & dtTemp.Rows(k)("AppID") & "'")
                                End If
                                Querypack = cCommon.ExecuteNonQueryWithQuerypack("DELETE EmpTransmittal WHERE ActivityID = '" & dtTemp.Rows(k)("ActivityID").ToString & "'")
                            Next

                        End If


                        ' ___ Delete bogus records
                        'For j = 0 To dtCallDate.Rows.Count - 1
                        'CallDate = dtCallDate.Rows(j)(0)
                        sb.Length = 0
                        sb.Append("SELECT et.ActivityID, et.EnrollerID, ept.AppID, ept.AltProductDataID ")
                        sb.Append("FROM ProjectReports..EmpTransmittal et ")
                        sb.Append("INNER JOIN UserManagement..Users u ON et.EnrollerID = u.UserID ")
                        sb.Append("LEFT JOIN EmpProductTransmittal ept on et.ActivityID = ept.ActivityID ")
                        'sb.Append("WHERE et.SupervisorApprovalDate IS NULL AND u.Role NOT IN ('ENROLLER', 'SUPERVISOR') AND dbo.ufn_IsDateEqual(et.CallStartTime, '" & CallDate & "') = 1")

                        sb.Append("WHERE et.SupervisorApprovalDate IS NULL AND u.Role NOT IN ('ENROLLER', 'SUPERVISOR') AND ")
                        sb.Append("dbo.ufn_IsDateBetween(et.CallStartTime, '" & dtCallDate.Rows(0)(0) & "', '" & dtCallDate.Rows(dtCallDate.Rows.Count - 1)(0) & "') = 1 ")
                        sb.Append("ORDER BY et.CallStartTime")

                        Querypack = cCommon.GetDTWithQuerypack(sb.ToString)
                        dtTemp = Querypack.dt

                        For k = 0 To dtTemp.Rows.Count - 1
                            If Not IsDBNull(dtTemp.Rows(k)("AltProductDataID")) Then
                                Querypack = cCommon.ExecuteNonQueryWithQuerypack("DELETE Alt_ProductData WHERE AltProductDataID = " & dtTemp.Rows(k)("AltProductDataID"))
                            End If
                            If Not IsDBNull(dtTemp.Rows(k)("AppID")) Then
                                Querypack = cCommon.ExecuteNonQueryWithQuerypack("DELETE EmpProductTransmittal WHERE AppID = '" & dtTemp.Rows(k)("AppID") & "'")
                            End If
                            Querypack = cCommon.ExecuteNonQueryWithQuerypack("DELETE EmpTransmittal WHERE ActivityID = '" & dtTemp.Rows(k)("ActivityID").ToString & "'")
                        Next
                        'Next

                        ' ___ Determine whether any of the call records for the dates selected for update have not been supervisor-approved
                        Querypack = cCommon.GetDTWithQuerypack("SELECT ActivityID = CAST(ActivityID as varchar(36)), CallStartTime, SupervisorApprovalDate, EnrollerID, ClientID, EmpName = LastName + ', ' + FirstName FROM ProjectReports..EmpTransmittal WHERE SupervisorApprovalDate IS NULL AND LogicalDelete = 0 AND dbo.ufn_IsDateBetween(CallStartTime, '" & dtCallDate.Rows(0)(0) & "', '" & dtCallDate.Rows(dtCallDate.Rows.Count - 1)(0) & "') = 1 ORDER BY CallStartTime")
                        For k = 0 To Querypack.dt.Rows.Count - 1
                            'OKToProceed = False
                            MyResults.Success = False
                            CallDate = Querypack.dt.Rows(k)("CallStartTime")
                            If ApprovalError = Nothing Then
                                ApprovalError = "CallDate: " & CallDate.ToString("MM/dd/yyyy HH:mm:ss") & ", ActivityID: " & Querypack.dt.Rows(k)("ActivityID") & ", EnrollerID: " & Querypack.dt.Rows(k)("EnrollerID") & ", ClientID: " & Querypack.dt.Rows(k)("ClientID") & ", EmpName: " & Querypack.dt.Rows(k)("EmpName") & "." & Environment.NewLine
                                System.Diagnostics.Debug.Write("'" & Querypack.dt.Rows(k)("ActivityID") & "',")
                            Else
                                ApprovalError &= " " & "CallDate: " & CallDate.ToString("MM/dd/yyyy HH:mm:ss") & ", ActivityID: " & Querypack.dt.Rows(k)("ActivityID") & ", EnrollerID: " & Querypack.dt.Rows(k)("EnrollerID") & ", ClientID: " & Querypack.dt.Rows(k)("ClientID") & ", EmpName: " & Querypack.dt.Rows(k)("EmpName") & "." & Environment.NewLine
                                System.Diagnostics.Debug.Write("'" & Querypack.dt.Rows(k)("ActivityID") & "',")
                            End If
                        Next

                        If k > 0 Then
                            Dim Response As DialogResult
                            'Response = MessageBox.Show("Unsupervisor-approved records have been found. Would you like to update the tables anyway?" & Environment.NewLine & Environment.NewLine & ApprovalError, "Unsupervisor-approved records", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2)
                            Response = MessageBox.Show("Unsupervisor-approved records have been found. Would you like to update the tables anyway?", "Unsupervisor-approved records", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2)
                            If Response <> DialogResult.Yes Then
                                OKToProceed = False
                            End If
                        End If

                        If Not OKToProceed Then
                            MyResults.Message = "Unsupervisor approved record(s): " & ApprovalError
                            Return MyResults
                        End If
                    End If

                    ' // *** At this point all of the tests have successfully been passed and the table updates are about to begin ***

                    ' ___ If update is successful, bail. Otherwise, write error to Rpt_Control and stay in loop if under maxiumum iterations
                    If OKToProceed Then
                        'Results = UpdateReportTables(dtClient, dtCallDate, NewAddDate)
                        Results = UpdateReportTables(dtCallDate, NewAddDate)
                        If Results.Success Then
                            ThisUpdateSuccessfulInd = True
                            Exit For
                        Else
                            MyResults.Success = False
                            MyResults.Message = Results.Message
                            Results = cCommon.GetNewRecordID("Rpt_Control", "RecID")
                            If Results.Success Then
                                RecID = Results.Value
                            Else
                                OKToProceed = False
                                ErrorMessage = "Try #" & i.ToString & " Sql: " & Sql & ": " & Results.Message
                                FailReason = Results.Message
                            End If

                            Querypack = cCommon.ExecuteNonQueryWithQuerypack("INSERT INTO Rpt_Control (RecID, AddDate, TryNum, SuccessInd, ErrorMessage) VALUES (" & RecID & ", '" & NewAddDate & "', " & i.ToString & ", 0, 'Error #101a: " & cCommon.StrOutHandler(cErrDetail & " " & MyResults.Message & " " & cLastSql, False, Common.StringTreatEnum.SecApost) & "')")

                            If Not Querypack.Success Then
                                OKToProceed = False
                                ErrorMessage = "Try #" & i.ToString & " Sql: " & Sql & ": " & Querypack.TechErrMsg
                                FailReason = Querypack.TechErrMsg
                            End If

                        End If
                    End If

                    If Not OKToProceed Then
                        Results = cCommon.GetNewRecordID("Rpt_Control", "RecID")
                        RecID = Results.Value
                        Querypack = cCommon.ExecuteNonQueryWithQuerypack("INSERT INTO Rpt_Control (RecID, AddDate, TryNum, SuccessInd, ErrorMessage) VALUES (" & RecID & ", '" & NewAddDate & "', " & i.ToString & ", 0, 'Error #101b: " & cCommon.StrOutHandler(cErrDetail & " " & cLastSql & " " & FailReason, False, Common.StringTreatEnum.SecApost) & "')")
                    End If

                Catch ex As Exception
                    Try
                        Results = cCommon.GetNewRecordID("Rpt_Control", "RecID")
                        RecID = Results.Value
                        'Querypack = cCommon.ExecuteNonQueryWithQuerypack("INSERT INTO Rpt_Control (RecID, AddDate, SuccessInd) VALUES (" & RecID & ", '" & NewAddDate & "', 0, 'End Try #" & i.ToString & "  loop " & ex.Message)
                        Querypack = cCommon.ExecuteNonQueryWithQuerypack("INSERT INTO Rpt_Control (RecID, AddDate, TryNum, SuccessInd, ErrorMessage) VALUES (" & RecID & ", '" & NewAddDate & "', " & i.ToString & ", 0, 'Error #101c: " & cCommon.StrOutHandler(cErrDetail & " " & cLastSql & " " & ex.Message, False, Common.StringTreatEnum.SecApost) & "')")
                    Catch
                    End Try
                End Try

            Next

            If ThisUpdateSuccessfulInd Then
                Try
                    Results = cCommon.GetNewRecordID("Rpt_Control", "RecID")
                    RecID = Results.Value
                    'Querypack = cCommon.ExecuteNonQueryWithQuerypack("INSERT INTO Rpt_Control (RecID, AddDate, SuccessInd) VALUES (" & RecID & ", '" & NewAddDate & "', 1)")
                    Querypack = cCommon.ExecuteNonQueryWithQuerypack("INSERT INTO Rpt_Control (RecID, AddDate, TryNum, SuccessInd) VALUES (" & RecID & ", '" & NewAddDate & "', " & i.ToString & ", 1)")
                Catch
                End Try
            End If

            If ThisUpdateSuccessfulInd Then
                MyResults.Success = True
            Else
                MyResults.Success = False
                MyResults.Message = Results.Message
            End If

            Return MyResults

            MyResults.Success = True
            Return MyResults

        Catch ex As Exception
            MyResults.Success = False
            MyResults.Message = "Main.HandleReportTablesUpdate #101d: " & cErrDetail & " " & cLastSql & " " & ex.Message
            Return MyResults
        End Try
    End Function

    Private Sub ResetVariables(ByVal i As Integer, ByRef UpdateType As UpdateTypeEnum, ByRef RptControlIsPopulated As Boolean, _
        ByRef TriggerDate As DateTime, ByRef CrashCallDate As DateTime, ByRef OKToProceed As Boolean, _
        ByRef ErrorMessage As String, ByRef FirstTimeThrough As Boolean)

        UpdateType = UpdateTypeEnum.Undetermined
        RptControlIsPopulated = False
        TriggerDate = "1/1/2000"
        CrashCallDate = "1/1/2000"
        OKToProceed = True
        ErrorMessage = Nothing

        ' ___ First time through
        If i = 1 Then
            FirstTimeThrough = True
        Else
            FirstTimeThrough = False
        End If
    End Sub

    Private Sub IsRpt_ControlPopulated(ByVal i As Integer, ByRef RptControlIsPopulated As Boolean, ByRef OKToProceed As Boolean, ByRef ErrorMessage As String, ByRef FailReason As String)
        Dim Sql As String
        Dim Querypack As QueryPack

        ' ___ Are there any records in Rpt_Control?
        Sql = "SELECT Count (*) FROM Rpt_Control"
        Querypack = cCommon.GetDTWithQuerypack(Sql)
        If Querypack.Success Then
            If Querypack.dt.Rows(0)(0) > 0 Then
                RptControlIsPopulated = True
            End If
        Else
            OKToProceed = False
            ErrorMessage = "IsRpt_ControlPopulated: Try #" & i.ToString & " Sql: " & Sql & ": " & Querypack.TechErrMsg
            FailReason = Querypack.TechErrMsg
        End If
    End Sub

    Private Sub GetUpdateTypeAndTriggerDateIfFirstTimeThrough(ByRef RptControlIsPopulated As Boolean, ByRef UpdateType As UpdateTypeEnum, _
        ByRef TriggerDate As DateTime, ByRef OKToProceed As Boolean, ByRef ErrorMessage As String, ByRef FailReason As String)

        Dim Sql As String
        Dim Querypack As QueryPack

        If RptControlIsPopulated Then

            ' ___ Was the update from the very last session successful?
            'Sql = "SELECT * FROM Rpt_Control WHERE AddDate = (SELECT Max(AddDate) FROM Rpt_Control)"
            Sql = "SELECT * FROM Rpt_Control rc WHERE AddDate = (SELECT Max(AddDate) FROM Rpt_Control) AND TryNum = (SELECT Max(TryNum) FROM Rpt_Control WHERE AddDate = rc.AddDate)"

            Querypack = cCommon.GetDTWithQuerypack(Sql)
            If Querypack.Success Then
                If Querypack.dt.Rows(0)("SuccessInd") Then
                    UpdateType = UpdateTypeEnum.SuccessfulUpdateVeryLastSession
                    TriggerDate = Querypack.dt.Rows(0)("AddDate")
                End If
            Else
                OKToProceed = False
                ErrorMessage = "GetUpdateTypeAndTriggerDateIfFirstTimeThrough: Sql: " & Sql & ": " & Querypack.TechErrMsg
                FailReason = Querypack.TechErrMsg
            End If

        Else

            ' ___ Rpt_Control is empty
            UpdateType = UpdateTypeEnum.RptControlIsEmpty
        End If

    End Sub

    Private Sub GetUpdateTypeAndCrashCallDateIfCallDatesUpdatedLastIteration(ByVal i As Integer, ByRef NewAddDate As DateTime, ByRef UpdateType As UpdateTypeEnum, _
        ByRef CrashCallDate As DateTime, ByRef OKToProceed As Boolean, ByRef ErrorMessage As String, ByRef FailReason As String)

        ' updatetype and crashcalldate

        Dim Sql As String
        Dim Querypack As QueryPack
        Dim IsSameAddDate As Boolean

        ' ___ Were any CallDates successfully updated during a previous iteration?
        'Sql = "SELECT Max(CallDate) FROM Rpt_CallHistory WHERE AddDate = '" & NewAddDate & "'"
        'Querypack = cCommon.GetDTWithQuerypack(Sql)
        'If Querypack.Success Then
        '    If Querypack.dt.Rows.Count > 0 Then
        '        UpdateType = UpdateTypeEnum.SuccessfulUpdatePriorIteration
        '        CrashCallDate = Querypack.dt.Rows(0)(0)
        '    End If
        'Else
        '    OKToProceed = False
        '    ErrorMessage = "Try #" & i.ToString & " Sql: " & Sql & ": " & Querypack.TechErrMsg
        '    FailReason = Querypack.TechErrMsg
        'End If

        ' Modified 6/2/2010
        'Sql = "SELECT Count (*) FROM Rpt_CallHistory WHERE AddDate = '" & NewAddDate & "'"
        'Querypack = cCommon.GetDTWithQuerypack(Sql)
        'If Querypack.Success Then
        '    If Not IsDBNull(Querypack.dt.rows(0)(0)) Then
        '        Sql = "SELECT Max(CallDate) FROM Rpt_CallHistory WHERE AddDate = '" & NewAddDate & "'"
        '        Querypack = cCommon.GetDTWithQuerypack(Sql)
        '        If Querypack.Success Then
        '            UpdateType = UpdateTypeEnum.SuccessfulUpdatePriorIteration
        '            CrashCallDate = Querypack.dt.Rows(0)(0)
        '        Else
        '            OKToProceed = False
        '            ErrorMessage = "Try #" & i.ToString & " Sql: " & Sql & ": " & Querypack.TechErrMsg
        '            FailReason = Querypack.TechErrMsg
        '        End If
        '    End If
        'Else
        '    OKToProceed = False
        '    ErrorMessage = "Try #" & i.ToString & " Sql: " & Sql & ": " & Querypack.TechErrMsg
        '    FailReason = Querypack.TechErrMsg
        'End If

        Sql = "SELECT Count (*) FROM Rpt_CallHistory WHERE AddDate = '" & NewAddDate & "'"
        Querypack = cCommon.GetDTWithQuerypack(Sql)
        If Querypack.Success Then
            If Querypack.dt.rows(0)(0) > 0 Then
                IsSameAddDate = True
            End If
        End If

        If Querypack.Success AndAlso IsSameAddDate Then
            Sql = "SELECT Max(CallDate) FROM Rpt_CallHistory WHERE AddDate = '" & NewAddDate & "'"
            Querypack = cCommon.GetDTWithQuerypack(Sql)
            If Querypack.Success Then
                UpdateType = UpdateTypeEnum.SuccessfulUpdatePriorIteration
                CrashCallDate = Querypack.dt.Rows(0)(0)
            End If
        End If

        If Not Querypack.Success Then
            OKToProceed = False
            ErrorMessage = "GetUpdateTypeAndCrashCallDateIfCallDatesUpdatedLastIteration: Try #" & i.ToString & " Sql: " & Sql & ": " & Querypack.TechErrMsg
            FailReason = Querypack.TechErrMsg
        End If
    End Sub


    Private Sub WasThereEverASuccessfulUpdate(ByVal i As Integer, ByRef UpdateType As UpdateTypeEnum, ByRef TriggerDate As DateTime, ByRef OKToProceed As Boolean, ByRef ErrorMessage As String, ByRef FailReason As String)
        Dim Sql As String
        Dim Querypack As QueryPack

        ' ___ Has there ever been a successul update?
        Sql = "SELECT Max(AddDate) FROM Rpt_Control WHERE SuccessInd = 1"
        Querypack = cCommon.GetDTWithQuerypack(Sql)

        If Querypack.Success Then
            If IsDBNull(Querypack.dt.Rows(0)(0)) Then
                UpdateType = UpdateTypeEnum.NeverBeenASuccessfulUpdate
            Else
                UpdateType = UpdateTypeEnum.SuccessfulUpdateMoreThanOneSessionAgo
                TriggerDate = Querypack.dt.Rows(0)(0)
            End If
        Else
            OKToProceed = False
            ErrorMessage = "WasThereEverASuccessfulUpdateTry #" & i.ToString & " Sql: " & Sql & ": " & Querypack.TechErrMsg
            FailReason = Querypack.TechErrMsg
        End If
    End Sub

    Private Function CleanUpHistoryTables(ByVal Sql As String) As Results
        Dim MyResults As New Results
        Dim Querypack As QueryPack

        Try

            Sql = Replace(Sql, "[tablename]", "Rpt_CallHistory")
            Querypack = cCommon.ExecuteNonQueryWithQuerypack(Sql)
            If Not Querypack.Success Then
                MyResults.Success = False
                MyResults.Message = Querypack.TechErrMsg
                Return MyResults
            End If

            Sql = Replace(Sql, "[tablename]", "Rpt_ProductHistory")
            Querypack = cCommon.ExecuteNonQueryWithQuerypack(Sql)
            If Not Querypack.Success Then
                MyResults.Success = False
                MyResults.Message = Querypack.TechErrMsg
                Return MyResults
            End If

            MyResults.Success = True
            Return MyResults

        Catch ex As Exception
            MyResults.Success = False
            MyResults.Message = "CleanUpHistoryTables: " & Replace(ex.Message, "'", "''")
            Return MyResults
        End Try
    End Function

    Private Function PerformCallDateDTTreatment(ByRef dtCallDate As DataTable, ByVal UpdateType As UpdateTypeEnum, ByVal FirstCallDate As Date, ByVal TriggerDate As DateTime, ByVal CrashCallDate As DateTime) As Results
        Dim i As Integer
        Dim MyResults As New Results
        Dim Sql As New System.Text.StringBuilder
        Dim Querypack As QueryPack
        Dim dt As DataTable
        Dim dr As DataRow
        Dim Results As Results

        Try

            If UpdateType = UpdateTypeEnum.SuccessfulUpdatePriorIteration Then

                ' ___ We are going to retain all dates written prior to the CrashCallDate. Delete dates up to and including the CrashCallDate from dtCallDate.
                dt = New DataTable
                dt.Columns.Add(New DataColumn("CallDate", GetType(System.String)))
                For i = 0 To dtCallDate.Rows.Count - 1
                    If cCommon.DateCompare(dtCallDate.Rows(i)(0), CrashCallDate, True) > -1 Then
                        dr = dt.NewRow
                        dr(0) = dtCallDate.Rows(i)(0)
                        dt.Rows.Add(dr)
                    End If
                Next
                dtCallDate = dt

            Else
                ' ___ Get a list of call dates from EmpTransmittal and EmpProductTransmittal whose ChangeDate falls after the AddDate in the Rpt_CallHistory and Rpt_CallProduct tables
                'Sql.Append("SELECT DISTINCT Convert(varchar, et.CallStartTime, 101) FROM EmpTransmittal et ")
                'Sql.Append("INNER JOIN EmpProductTransmittal ept on et.ActivityID = ept.ActivityID ")
                'Sql.Append("WHERE (et.CallStartTime >= '" & FirstCallDate & "') AND ")
                'Sql.Append("(et.ChangeDate > '" & TriggerDate & "'  OR ept.ChangeDate > '" & TriggerDate & "') ")
                'Sql.Append("ORDER BY Convert(varchar, et.CallStartTime, 101) ")




                Sql.Append("SELECT DISTINCT Convert(varchar, et.CallStartTime, 101) FROM EmpTransmittal et ")
                Sql.Append("LEFT JOIN EmpProductTransmittal ept on et.ActivityID = ept.ActivityID ")
                Sql.Append("WHERE et.CallStartTime >= '" & FirstCallDate & "' AND ")
                'Sql.Append("(dbo.ufn_DateCompare(et.ChangeDate, '" & TriggerDate & "', 1) > -1 OR (ept.ChangeDate IS NOT NULL AND  dbo.ufn_DateCompare(ept.ChangeDate, '" & TriggerDate & "', 1) > -1)) ")
                Sql.Append("(dbo.ufn_DateCompare(et.ChangeDate, '" & TriggerDate & "', 0) > -1 OR (ept.ChangeDate IS NOT NULL AND  dbo.ufn_DateCompare(ept.ChangeDate, '" & TriggerDate & "', 0) > -1)) ")
                Sql.Append("ORDER BY Convert(varchar, et.CallStartTime, 101) ")

                Querypack = cCommon.GetDTWithQuerypack(Sql.ToString)
                If Not Querypack.Success Then
                    MyResults.Success = False
                    MyResults.Message = "PerformCallDateDTTreatment Error #100a: " & Querypack.TechErrMsg
                    Return MyResults
                End If
                dtCallDate = Querypack.dt
            End If

            ' ___ Sql for returning all records -- Used for diagnostics only
            Sql.Length = 0
            Sql.Append("SELECT Convert(varchar, et.CallStartTime, 101), ")
            Sql.Append("et.*, ept.* FROM EmpTransmittal et ")
            Sql.Append("LEFT JOIN EmpProductTransmittal ept on et.ActivityID = ept.ActivityID ")
            Sql.Append("WHERE et.CallStartTime >= '" & FirstCallDate & "' AND ")
            Sql.Append("(dbo.ufn_DateCompare(et.ChangeDate, '" & TriggerDate & "', 1) > -1 OR (ept.ChangeDate IS NOT NULL AND  dbo.ufn_DateCompare(ept.ChangeDate, '" & TriggerDate & "', 1) > -1)) ")



            ' ___ Exclude today
            If Not cEnviro.IncludeTodayInUpdate Then
                Sql.Append("AND dbo.ufn_IsDateEqual(et.CallStartTime, getDate()) <> 1 ")
            End If

            Sql.Append("ORDER BY Convert(varchar, et.CallStartTime, 101) ")
            System.Diagnostics.Debug.Write(Sql.ToString)

            If dtCallDate.Rows.Count > 0 Then
                Results = PerformCallDateDTTreatment2(dtCallDate)
                If Not Results.Success Then
                    MyResults.Success = False
                    MyResults.Message = "PerformCallDateDTTreatment Error #100b: " & Results.Message
                    Return MyResults
                End If
            End If

            If cCallType = CallTypeEnum.UpdateDates Then
                If dtCallDate.Rows.Count = 0 Then
                    MyResults.Message = "No dates selected."
                Else
                    For i = 0 To dtCallDate.Rows.Count - 1
                        If i = 0 Then
                            MyResults.Message = dtCallDate.Rows(i)(0)
                        Else
                            MyResults.Message &= ", " & dtCallDate.Rows(i)(0)
                        End If
                    Next
                End If
            End If

            MyResults.Success = True
            Return MyResults

        Catch ex As Exception
            MyResults.Success = False
            MyResults.Message = "Main.PerformCallDateDTTreatment #100c: " & Replace(ex.Message, "'", "''")
            Return MyResults
        End Try
    End Function

    Private Function PerformCallDateDTTreatment2(ByRef dtCallDate As DataTable) As Results
        Dim i As Integer
        Dim MyResults As New Results
        Dim ReportConfig As ReportConfig
        Dim ReportDate As DateTime
        Dim ReportYear As Integer
        Dim CallYear As Integer
        Dim FirstBusinessDate As DateTime
        Dim CallDate As String
        Dim dt As DataTable

        Try

            ' ___ Condition the CallDate table
            ReportConfig = New ReportConfig(-1)
            If dtCallDate.Rows.Count > 0 Then
                dt = dtCallDate.Clone
            End If

            ReportDate = CType(ReportConfig.ReportDate, System.DateTime)
            ReportYear = CType(ReportDate.ToString("yyyy"), System.Int64)

            'If Not cEnviro.IncludeTodayInUpdate Then
            '    If cCommon.DateCompare(dtCallDate.Rows(dtCallDate.Rows.Count - 1)(0), cCommon.GetServerDateTime, True) = 0 Then
            '        dtCallDate.Rows.RemoveAt(dtCallDate.Rows.Count - 1)
            '    End If
            'End If

            ' ___ Remove today's date in accordance with application setting.
            If Not cEnviro.IncludeTodayInUpdate Then
                For i = 0 To dtCallDate.Rows.Count - 1
                    CallDate = dtCallDate.Rows(i)(0)
                    If Not cCommon.DateCompare(CType(CallDate, System.DateTime), ReportDate, True) = 0 Then
                        dt.Rows.Add(dt.NewRow)
                        dt.Rows(dt.Rows.Count - 1)(0) = CallDate
                    End If
                Next
            End If

            ' ___ Clear out dtCallDate. We are going to re-use it.
            dtCallDate.Rows.Clear()

            ' // If update for previous year is not allowed, remove previous year dates
            ' // from table unless report date is the first business date of the new year (Jan 2).
            ' // Also, remove any records for years previous to last year.

            ' ___ Get the first business date of the new year, either Jan 2 or a different date if Jan 2 falls on a weekend
            FirstBusinessDate = cCommon.GetFirstBusinessDateOfYear(ReportYear)

            For i = 0 To dt.Rows.Count - 1
                CallDate = dt.Rows(i)(0)
                CallYear = CType(cCommon.Right(CType(CallDate, System.String), 4), System.Int64)

                If CallYear = ReportYear Then
                    dtCallDate.Rows.Add(dtCallDate.NewRow)
                    dtCallDate.Rows(dtCallDate.Rows.Count - 1)(0) = CallDate
                ElseIf CallYear = ReportYear - 1 Then
                    If cEnviro.AllowTableUpdatePreviousYear Then
                        dtCallDate.Rows.Add(dtCallDate.NewRow)
                        dtCallDate.Rows(dtCallDate.Rows.Count - 1)(0) = CallDate
                    Else
                        If cCommon.DateCompare(CallDate, FirstBusinessDate, False) = 0 Then
                            dtCallDate.Rows.Add(dtCallDate.NewRow)
                            dtCallDate.Rows(dtCallDate.Rows.Count - 1)(0) = CallDate
                        End If
                    End If
                Else
                    ' No other year entry allowed
                End If
            Next

            MyResults.Success = True
            Return MyResults

        Catch ex As Exception
            MyResults.Success = False
            MyResults.Message = "PerformCallDateDTTreatment2: Error #104: " & ex.Message
            Return MyResults
        End Try
    End Function
#End Region

#Region " UpdateReportTables control "
    Private Function UpdateReportTables(ByRef dtCallDate As DataTable, ByVal AddDate As DateTime) As Results
        Dim i, j As Integer
        Dim MyResults As New Results
        Dim Results As Results
        Dim CallDate As Date
        Dim ClientID As String
        Dim Querypack As QueryPack
        Dim dtClient As DataTable
        Dim Sql As New System.Text.StringBuilder

        Try

            ' ___ Output
            e.Message = "Main.UpdateReportTables"
            RaiseEvent NotifyForm(e)

            ' ___ Set up HandleClientForThisDate
            For i = 0 To dtCallDate.Rows.Count - 1
                CallDate = dtCallDate.Rows(i)(0)


                ' ___ Build the client table for this date
                Sql.Length = 0
                Sql.Append("SELECT ClusterID FROM Excel_Cluster ")
                Sql.Append("WHERE ProdRptStatusInd = 1 AND ")
                Sql.Append("StartDate IS NOT NULL AND dbo.ufn_DateCompare(StartDate, '" & CallDate & "', 1) < 1 AND ")
                Sql.Append("(EndDate IS NULL OR dbo.ufn_DateCompare(EndDate, '" & CallDate & "', 1) > -1) ")
                Sql.Append("ORDER BY ClusterID")
                Querypack = cCommon.GetDTWithQuerypack(Sql.ToString)

                'QueryPack = cCommon.GetDTWithQuerypack("SELECT SegmentID FROM Excel_SegmentConfigure WHERE SegmentType = 'CLIENT' ORDER BY SegmentID")

                If Not Querypack.Success Then
                    MyResults.Success = False
                    MyResults.Message = Querypack.TechErrMsg
                    Return MyResults
                End If
                dtClient = Querypack.dt


                For j = 0 To dtClient.Rows.Count - 1
                    ClientID = dtClient.Rows(j)(0)

                    Results = Update_Rpt_CallHistory_ThisClient_ThisDate(CallDate, ClientID, AddDate)
                    If Not Results.Success Then
                        MyResults.Success = False
                        MyResults.Message = Results.Message
                        Return MyResults
                    End If

                    Results = Update_Rpt_ProductHistory_ThisClient_ThisDate(CallDate, ClientID, AddDate)
                    If Not Results.Success Then
                        MyResults.Success = False
                        MyResults.Message = Results.Message
                        Return MyResults
                    End If

                Next

                ' ___ Test for client completion
                If ClientID.ToUpper <> dtClient.Rows(dtClient.Rows.Count - 1)(0).ToUpper Then
                    MyResults.Success = False
                    MyResults.Message = "Main.UpdateReportTables #102a: Failed client completion test "
                    Return MyResults
                End If

            Next

            ' ___ Test for date completion
            If dtCallDate.Rows.Count > 0 Then
                If CallDate <> dtCallDate.Rows(dtCallDate.Rows.Count - 1)(0) Then
                    MyResults.Success = False
                    MyResults.Message = "Main.UpdateReportTables #101: Failed date completion test "
                    Return MyResults
                End If
            End If

            MyResults.Success = True
            Return MyResults

        Catch ex As Exception
            MyResults.Success = False
            MyResults.Message = "UpdateReportTables: Error #102b: " & Replace(ex.Message, "'", "''")
            Return MyResults
        End Try
    End Function
#End Region

#Region " Rpt_CallHistory"
    Private Function Update_Rpt_CallHistory_ThisClient_ThisDate(ByVal CallDate As Date, ByVal ClientID As String, ByVal AddDate As DateTime) As Results
        Dim MyResults As New Results
        Dim CmdPack As CmdPack
        Dim Sql As String = String.Empty
        Dim ErrDetail As String

        Try

            ErrDetail = "CallDate: " & CallDate.ToString("MM/dd/yyyy") & ", ClientID: " & ClientID & ", AddDate: " & AddDate.ToString

            ' ___ Output
            cErrDetail = "Main.Update_Rpt_CallHistory_ThisClient_ThisDate: " & ErrDetail
            e.Message = "Main.Update_Rpt_CallHistory_ThisClient_ThisDate: " & ErrDetail
            RaiseEvent NotifyForm(e)

            ' ___ Delete the records that are about to be replaced
            'Sql = "DELETE Rpt_CallHistory WHERE ClientID = '" & ClientID & "' AND dbo.ufn_IsDateEqual(CallDate, '" & CallDate & "') = 1"
            'cLastSql = Sql
            'Querypack = cCommon.ExecuteNonQueryWithQuerypack(Sql)
            'If Not Querypack.Success Then
            ' MyResults.Success = False
            ' MyResults.Message = "Update_Rpt_CallHistory_ThisClient_ThisDate #107a: " & Querypack.TechErrMsg
            ' Return MyResults
            ' End If

            CmdPack = New CmdPack("usp_RptBuildClientSegment101", CommandType.StoredProcedure, cEnviro)
            CmdPack.AddParameter(SqlDbType.DateTime, "@CallDate", CallDate, ParameterDirection.Input)
            CmdPack.AddParameter(SqlDbType.DateTime, "@AddDate", AddDate, ParameterDirection.Input)
            CmdPack.AddParameter(SqlDbType.VarChar, "@ClientID", ClientID, ParameterDirection.Input, 20)
            CmdPack.AddParameter(SqlDbType.VarChar, "@RetVal", "''", ParameterDirection.Output, 8000)
            CmdPack.Execute()
            If Not CmdPack.Success Then
                MyResults.Success = False
                MyResults.Message = "Main.Update_Rpt_CallHistory_ThisClient_ThisDate #107b: " & CmdPack.TechErrMsg
                Return MyResults
            End If

            If Not CmdPack.ParameterColl("@RetVal") = "0|Success" Then
                MyResults.Success = False
                MyResults.Message = "Update_Rpt_CallHistory_ThisClient_ThisDate #107c: " & CmdPack.ParameterColl("@RetVal")
                Return MyResults
            End If

            MyResults.Success = True
            Return MyResults

        Catch ex As Exception
            MyResults.Success = False
            MyResults.Message = "Update_Rpt_CallHistory_ThisClient_ThisDate #107d: " & Replace(ex.Message, "'", "''")
            Return MyResults
        End Try
    End Function
#End Region

#Region " Rpt_ProductHistory "
    Private Function Update_Rpt_ProductHistory_ThisClient_ThisDate(ByVal CallDate As Date, ByVal ClientID As String, ByVal AddDate As DateTime) As Results
        Dim i As Integer
        Dim Querypack As QueryPack
        Dim MyResults As New Results
        Dim dtProduct As DataTable
        Dim Results As Results
        Dim ProductID As String
        Dim ErrDetail As String
        Dim CompositeClientInd As Boolean
        Dim SubClientID As String
        Dim SubProductID As String
        Dim Sql As String

        Try

            ErrDetail = "CallDate: " & CallDate.ToString("MM/dd/yyyy") & ", ClientID: " & ClientID & "  "

            ' ___ Output
            cErrDetail = "Main.Update_Rpt_ProductHistory_ThisClient_ThisDate: " & cErrDetail
            e.Message = "Main.Update_Rpt_ProductHistory_ThisClient_ThisDate: " & ErrDetail
            RaiseEvent NotifyForm(e)

            ' ___ Composite client?
            If InStr(ClientID, "|") > 0 Then
                CompositeClientInd = True
            End If

            ' ___ Build the product table
            Results = GetProductTable(dtProduct, ClientID)
            If Not Results.Success Then
                MyResults.Success = False
                MyResults.Message = Results.Message
                Return MyResults
            End If

            ' ___ Delete the records that are about to be replaced
            'Querypack = cCommon.ExecuteNonQueryWithQuerypack("DELETE Rpt_ProductHistory WHERE ClientID = '" & ClientID & "' AND dbo.ufn_IsDateEqual(CallDate, '" & CallDate & "') = 1")

            Sql = "DELETE Rpt_ProductHistory WHERE " & GetClientWhereSelect(ClientID) & " dbo.ufn_IsDateEqual(CallDate, '" & CallDate & "') = 1"
            cLastSql = Sql
            Querypack = cCommon.ExecuteNonQueryWithQuerypack(Sql)
            If Not Querypack.Success Then
                MyResults.Success = False
                MyResults.Message = "Update_Rpt_ProductHistory_ThisClient_ThisDate #108a : " & ErrDetail & Querypack.TechErrMsg
                Return MyResults
            End If

            ' ___ Loop through the product table for this client for this date
            For i = 0 To dtProduct.Rows.Count - 1
                ProductID = dtProduct.Rows(i)(0)
                If CompositeClientInd Then
                    SubClientID = ProductID.Substring(0, InStr(ProductID, "~") - 1)
                    SubProductID = cCommon.Right(ProductID, ProductID.Length - InStr(ProductID, "~"))
                Else
                    SubClientID = ClientID
                    SubProductID = dtProduct.Rows(i)(0)
                End If
                Results = Update_Rpt_ProductHistory_ThisClient_ThisDate_ThisProduct(CallDate, SubClientID, SubProductID, AddDate)
                If Not Results.Success Then
                    MyResults.Success = False
                    MyResults.Message = Results.Message
                    Return MyResults
                End If
            Next

            ' ___ Test for product completion
            If dtProduct.Rows.Count > 0 Then
                If ProductID <> dtProduct.Rows(dtProduct.Rows.Count - 1)(0) Then
                    MyResults.Success = False
                    MyResults.Message = "Update_Rpt_ProductHistory_ThisClient_ThisDate #108b :  Failed product completion test " & ErrDetail
                    Return MyResults
                End If
            End If

            MyResults.Success = True
            Return MyResults

        Catch ex As Exception
            MyResults.Success = False
            MyResults.Message = "Update_Rpt_ProductHistory_ThisClient_ThisDate #108c : " & ErrDetail & Replace(ex.Message, "'", "''")
            Return MyResults
        End Try
    End Function

    Private Function GetClientWhereSelect(ByVal ClientID As String) As String
        Dim i As Integer
        Dim Results As String
        Dim Box As String()

        If InStr(ClientID, "|") = 0 Then
            Results = " ClientID = '" & ClientID & "' AND "
        Else
            Box = Split(ClientID, "|")
            For i = 0 To Box.GetUpperBound(0)
                Box(i) = "ClientID = '" & Box(i) & "'"
            Next

            Results = " ("
            For i = 0 To Box.GetUpperBound(0)
                If i < Box.GetUpperBound(0) Then
                    Results &= Box(i) & " OR "
                Else
                    Results &= Box(i)
                End If
            Next
            Results &= ") AND "
        End If

        Return Results
    End Function

    Private Function Update_Rpt_ProductHistory_ThisClient_ThisDate_ThisProduct(ByVal CallDate As Date, ByVal ClientID As String, ByVal ProductID As String, ByVal AddDate As DateTime) As Results
        Dim i As Integer
        Dim MyResults As New Results
        Dim Results As Results
        Dim Querypack As QueryPack
        Dim dtFieldName As DataTable
        Dim FieldName As String
        Dim ProductType As ProductTypeEnum
        Dim ErrDetail As String
        Dim Sql As String = String.Empty

        Try

            ErrDetail = "CallDate: " & CallDate.ToString("MM/dd/yyyy") & ", ClientID: " & ClientID & ", ProductID: " & ProductID & "  "

            ' ___ Output
            cErrDetail = "Update_Rpt_ProductHistory_ThisClient_ThisDate_ThisProduct: " & ErrDetail
            e.Message = "Update_Rpt_ProductHistory_ThisClient_ThisDate_ThisProduct: " & ErrDetail
            RaiseEvent NotifyForm(e)

            ' ___ Is this a standard or extended/Aces special product?
            Sql = "SELECT Count (*) FROM ClientProduct_Extended WHERE ClientID =  '" & ClientID & "'  AND ClientProductID = '" & ProductID & "' AND dbo.ufn_IsDateBetween('" & CallDate & "', StartDate, EndDate) = 1"
            cLastSql = Sql
            Querypack = cCommon.GetDTWithQuerypack(Sql)
            If Not Querypack.Success Then
                MyResults.Success = False
                MyResults.Message = "Main.Update_Rpt_ProductHistory_ThisClient_ThisDate_ThisProduct #100 : " & ErrDetail & Querypack.TechErrMsg
                Return MyResults
            End If
            If Querypack.dt.rows(0)(0) = 0 Then
                ProductType = ProductTypeEnum.Standard
            Else
                ProductType = ProductTypeEnum.Extended
            End If

            ' ___ FieldName table
            Sql = "Select * FROM dbo.ufn_GetTableFromList('" & ProductID & "')"
            cLastSql = Sql
            Querypack = cCommon.GetDTWithQuerypack(Sql)
            If Not Querypack.Success Then
                MyResults.Success = False
                MyResults.Message = "Update_Rpt_ProductHistory_ThisClient_ThisDate_ThisProduct #109a : " & ErrDetail & Querypack.TechErrMsg
                Return MyResults
            End If
            dtFieldName = Querypack.dt

            For i = 0 To dtFieldName.Rows.Count - 1
                FieldName = dtFieldName.Rows(i)(1)
                Results = Update_Rpt_ProductHistory_ThisClient_ThisDate_ThisProduct_ThisField(CallDate, ClientID, ProductID, ProductType, FieldName, AddDate)
                If Not Results.Success Then
                    MyResults.Success = False
                    MyResults.Message = Results.Message
                    Return MyResults
                End If
            Next

            MyResults.Success = True
            Return MyResults

        Catch ex As Exception
            MyResults.Success = False
            MyResults.Message = "Update_Rpt_ProductHistory_ThisClient_ThisDate_ThisProduct #109b : " & ErrDetail & Replace(ex.Message, "'", "''")
            Return MyResults
        End Try
    End Function

    Private Function Update_Rpt_ProductHistory_ThisClient_ThisDate_ThisProduct_ThisField(ByVal CallDate As Date, ByVal ClientID As String, ByVal ProductID As String, ByVal ProductType As ProductTypeEnum, ByVal FieldName As String, ByVal AddDate As DateTime) As Results
        Dim MyResults As New Results
        Dim SPName As String
        Dim CmdPack As CmdPack
        Dim Sql2 As String
        Dim ErrDetail As String

        Try

            Select Case ProductType
                Case ProductTypeEnum.Standard
                    SPName = "usp_RptBuildProductSegmentStandard1"
                Case ProductTypeEnum.Extended
                    SPName = "usp_RptBuildProductSegmentExtended1"
            End Select

            ErrDetail = "CallDate: " & CallDate.ToString("MM/dd/yyyy") & ", ClientID: " & ClientID & ", ProductID: " & ProductID & ", ProductType: " & ProductType.ToString & ", FieldName:  " & FieldName & "  "

            ' ___ Output
            cErrDetail = "Main.Update_Rpt_ProductHistory_ThisClient_ThisDate_ThisProduct_ThisField: " & ErrDetail & " Top"
            e.Message = "Main.Update_Rpt_ProductHistory_ThisClient_ThisDate_ThisProduct_ThisField: " & ErrDetail
            RaiseEvent NotifyForm(e)

            CmdPack = New CmdPack(SPName, CommandType.StoredProcedure, cEnviro)
            CmdPack.AddParameter(SqlDbType.DateTime, "@CallDate", CallDate, ParameterDirection.Input)
            CmdPack.AddParameter(SqlDbType.VarChar, "@ClientID", ClientID, ParameterDirection.Input, 20)
            CmdPack.AddParameter(SqlDbType.VarChar, "@ProductID", ProductID, ParameterDirection.Input, 20)
            CmdPack.AddParameter(SqlDbType.VarChar, "@FieldName", FieldName, ParameterDirection.Input, 20)
            CmdPack.AddParameter(SqlDbType.VarChar, "@Sql2", "''", ParameterDirection.Output, 5000)
            CmdPack.AddParameter(SqlDbType.VarChar, "@RetVal", "''", ParameterDirection.Output, 500)
            CmdPack.Execute()
            If Not CmdPack.Success Then
                MyResults.Success = False
                MyResults.Message = "Update_Rpt_ProductHistory_ThisClient_ThisDate_ThisProduct_ThisField #100: " & SPName & "  " & ErrDetail & CmdPack.TechErrMsg
                Return MyResults
            End If

            If IsDBNull(CmdPack.ParameterColl("@RetVal")) Then
                MyResults.Success = False
                MyResults.Message = "Main.Update_Rpt_ProductHistory_ThisClient_ThisDate_ThisProduct_ThisField #101:" & ErrDetail & "@RetVal is null"
                Return MyResults
            End If

            If Not CmdPack.ParameterColl("@RetVal") = ProductID Then
                MyResults.Success = False
                MyResults.Message = "Update_Rpt_ProductHistory_ThisClient_ThisDate_ThisProduct_ThisField #101:" & ErrDetail & CmdPack.ParameterColl("@RetVal")
                Return MyResults
            End If
            cErrDetail = "Main.Update_Rpt_ProductHistory_ThisClient_ThisDate_ThisProduct_ThisField: " & ErrDetail & " Between stored procedures"

            If CmdPack.ParameterColl("@Sql2").Length > cTempSql2Length Then
                cTempSql2Length = CmdPack.ParameterColl("@Sql2").Length
            End If

            Sql2 = CmdPack.ParameterColl("@Sql2")
            cLastSql = Sql2

            ' ___ Finalize the write with the compactor
            CmdPack = New CmdPack("usp_RptBuildProductCompactor", CommandType.StoredProcedure, cEnviro)
            CmdPack.AddParameter(SqlDbType.DateTime, "@CallDate", CallDate, ParameterDirection.Input)
            CmdPack.AddParameter(SqlDbType.DateTime, "@AddDate", AddDate, ParameterDirection.Input)
            CmdPack.AddParameter(SqlDbType.VarChar, "@ClientID", ClientID, ParameterDirection.Input, 20)
            CmdPack.AddParameter(SqlDbType.VarChar, "@ProductID", ProductID, ParameterDirection.Input, 20)
            CmdPack.AddParameter(SqlDbType.VarChar, "@FieldName", FieldName, ParameterDirection.Input, 20)
            CmdPack.AddParameter(SqlDbType.VarChar, "@Sql2", Sql2, ParameterDirection.Input, 5000)
            CmdPack.AddParameter(SqlDbType.VarChar, "@RetVal", "''", ParameterDirection.Output, 500)
            CmdPack.Execute()
            If Not CmdPack.Success Then
                MyResults.Success = False
                MyResults.Message = "Main.Update_Rpt_ProductHistory_ThisClient_ThisDate_ThisProduct_ThisField #102: " & ErrDetail & CmdPack.TechErrMsg
                Return MyResults
            End If
            If Not CmdPack.ParameterColl("@RetVal") = "0|Success" Then
                MyResults.Success = False
                MyResults.Message = "Main.Update_Rpt_ProductHistory_ThisClient_ThisDate_ThisProduct_ThisField #103: " & ErrDetail & CmdPack.ParameterColl("@Sql2")
                Return MyResults
            End If

            cErrDetail = "Main.Update_Rpt_ProductHistory_ThisClient_ThisDate_ThisProduct_ThisField: " & ErrDetail & " After stored procedures"

            MyResults.Success = True
            Return MyResults

        Catch ex As Exception
            MyResults.Success = False
            MyResults.Message = "Update_Rpt_ProductHistory_ThisClient_ThisDate_ThisProduct_ThisField #110a: " & ErrDetail & Replace(ex.Message, "'", "''")
            Return MyResults
        End Try
    End Function

    Private Function GetProductTable(ByRef dtProduct As DataTable, ByVal ClientID As String) As Results
        Dim i, j As Integer
        Dim Querypack As QueryPack
        Dim MyResults As New Results
        Dim dr As DataRow
        Dim ErrDetail As String
        Dim Box() As String
        Dim Sql As String = String.Empty

        Try

            ErrDetail = "ClientID: " & ClientID

            ' ___ Output
            cErrDetail = "Main.GetProductTable: " & ErrDetail
            e.Message = "Main.GetProductTable: " & ErrDetail
            RaiseEvent NotifyForm(e)

            Sql = "SELECT SegmentID, SegmentBlend FROM Excel_ClusterSegment WHERE ClusterID = '" & ClientID & "' ORDER By SegmentID"
            cLastSql = Sql
            Querypack = cCommon.GetDTWithQuerypack(Sql)
            If Not Querypack.Success Then
                MyResults.Success = False
                MyResults.Message = "Main.GetProductTable #111a: " & ErrDetail & Querypack.TechErrMsg
                Return MyResults
            End If
            dtProduct = Querypack.dt

            ' ___ Add product blends to the product table
            For i = 0 To dtProduct.Rows.Count - 1
                If Not IsDBNull(dtProduct.Rows(i)(1)) Then
                    Box = Split(dtProduct.Rows(i)(1), "~")
                    For j = 0 To Box.GetUpperBound(0)
                        If cCommon.RowItemFind(dtProduct, 0, Box(j)) = -1 Then
                            dr = dtProduct.NewRow
                            dr(0) = Box(j)
                            dtProduct.Rows.Add(dr)
                        End If
                    Next
                End If
            Next

            MyResults.Success = True
            Return MyResults

        Catch ex As Exception
            MyResults.Success = False
            MyResults.Message = "GetProductTable #111b: " & ErrDetail & Replace(ex.Message, "'", "''")
        End Try
    End Function
#End Region
End Class