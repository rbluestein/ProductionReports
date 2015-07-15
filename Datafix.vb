Public Class Datafix
    Private cCommon As New Common


    Public Sub AddBVIFLegal()
        Dim dt As DataTable
        Dim i, j As Integer
        Dim sb As New System.Text.StringBuilder
        Dim CallDate As DateTime

        Try

            dt = ExcelToDT("c:\apps\productionreports\tempdata\datafix2.xls", True)

            For i = 0 To dt.Rows.Count - 1
                For j = 0 To 1
                    sb.Length = 0
                    AddBVIFLegal2(i, dt, sb, CallDate)

                    Select Case j
                        Case 0
                            sb.Append("'Yes', '1.00', ")
                        Case 1
                            sb.Append("'AnnualPremium', '" & CType(dt.Rows(i)("AnnualPremium"), System.String) & ".00', ")
                    End Select

                    sb.Append("'" & cCommon.GetServerDateTime.ToString & "')")

                    cCommon.ExecuteNonQuerySPErrorHandlingWithQuerypack(sb.ToString)
                Next
            Next

        Catch ex As Exception
            Stop
        End Try

    End Sub

    Public Sub AddBVIFLegal2(ByVal i As Integer, ByRef dt As DataTable, ByRef sb As System.Text.StringBuilder, ByRef CallDate As Date)
        Dim EnrollerID As String
        Dim CallDateStr As String
        Dim ClientID As String
        Dim ProductID As String

        EnrollerID = "'" & dt.Rows(i)("LicensedEnroller") & "'"
        CallDate = dt.Rows(i)("CallStartTime")
        CallDateStr = "'" & CallDate.ToString("MM/dd/yyyy") & "'"
        ClientID = "'BureauVeritas'"
        ProductID = "'HYATTLEGAL'"

        sb.Append("INSERT INTO ProjectReports..Rpt_ProductHistory (EnrollerID, CallDate,  ClientID, ProductID, FieldName, FieldData, AddDate) ")
        sb.Append(" VALUES (")
        sb.Append(EnrollerID & ", ")
        sb.Append(CallDateStr & ", ")
        sb.Append(ClientID & ", ")
        sb.Append(ProductID & ", ")
    End Sub


    Public Sub PopulateCallbackMasterWithStantec()
        'Dim i, j As Integer
        'Dim dt As DataTable
        'Dim FullPath As String
        'Dim Sql As New System.Text.StringBuilder
        'Dim CallbackID As Integer
        'Dim TicketNumber As Integer
        'Dim dr As DataRow
        'Dim ItemArray As Object
        'Dim sb As New System.Text.StringBuilder
        'Dim ChangeDate As String
        'Dim LoggedInUserID As String = "jpenny"

        'Try

        '    FullPath = "C:\Apps\ExcelToTable\bin\Stantec.xls"

        '    dt = ExcelToDT(FullPath, True, "SELECT * FROM [AA.ECO$]")

        '    For i = 0 To dt.Rows.Count - 1
        '        ChangeDate = Common.GetServerDateTime

        '        dr = dt.Rows(i)
        '        ItemArray = dr.ItemArray

        '        'CallbackID = Common.GetNewRecordID("CallbackMaster", "CallbackID", 100001, 999999)
        '        'TicketNumber = Common.GetNewRecordID("CallbackMaster", "CallbackID", 100001, 999999)


        '        sb.Length = 0
        '        Sql.Append("INSERT INTO CallbackMaster ")
        '        Sql.Append("(")
        '        Sql.Append("CallbackID, CreationDate, TicketNumber, ClientID, State, ")
        '        Sql.Append("EmpLastName, EmpFirstName, EmpMI, EmpID, ")
        '        Sql.Append("PreferSpanishInd, OverflowAgentID, EnrollWinCode, ")
        '        Sql.Append("PriorityTagInd, CallPurposeCode, StatusCode, ")
        '        Sql.Append("Notes, NumEmployeeCalls, EnrollWinActivityID, EnrollWinStartDate, EnrollWinEndDate, LogicalDelete, ")
        '        Sql.Append("AddDate, ChangeDate")
        '        Sql.Append(") ")

        '        Sql.Append(" Values ")

        '        Sql.Append("(")
        '        Sql.Append(Common.NumOutHandler(CallbackID, False, False) & ", ")
        '        Sql.Append(Common.DateOutHandler(ChangeDate, False, True) & ", ")
        '        Sql.Append(Common.NumOutHandler(TicketNumber, False, False) & ", ")
        '        Sql.Append(Common.StrOutHandler("STANTEC", False, StringTreatEnum.SideQts) & ", ")
        '        Sql.Append(Common.StrOutHandler(dr("State"), False, StringTreatEnum.SideQts) & ",")

        '        Sql.Append(Common.StrOutHandler(dr("LastName"), False, StringTreatEnum.SideQts_SecApost) & ", ")
        '        Sql.Append(Common.StrOutHandler(dr("FirstName"), False, StringTreatEnum.SideQts_SecApost) & ", ")
        '        Sql.Append(Common.StrOutHandler("", False, StringTreatEnum.SideQts_SecApost) & ", ")
        '        Sql.Append(Common.StrOutHandler(dr("EmpID"), False, StringTreatEnum.SideQts) & ", ")

        '        Sql.Append("0, ")
        '        Sql.Append(Common.StrOutHandler(LoggedInUserID, False, StringTreatEnum.SideQts) & ", ")
        '        Sql.Append("'OE', ")

        '        Sql.Append("1, ")
        '        Sql.Append("'WE', ")
        '        Sql.Append("'CB', ")

        '        Sql.Append("'Auto add list 1', ")

        '        Sql.Append("0, ")

        '        Sql.Append("null, ")
        '        Sql.Append("null, ")
        '        Sql.Append(Common.DateOutHandler("12/10/2010", False, True) & ", ")

        '        Sql.Append("0 , ")
        '        Sql.Append(Common.DateOutHandler(ChangeDate, False, True) & ", ")
        '        Sql.Append(Common.DateOutHandler(ChangeDate, False, True))
        '        Sql.Append(") ")

        '        ' ___ Save CallbackMaster record
        '        Common.ExecuteNonQuery(Sql.ToString)

        '        ' ___ CallbackPhone
        '        For j = 2 To 1 Step -1
        '            PerformSave_CallbackPhone(dr, CallbackID, j)
        '        Next
        '    Next

        'Catch ex As Exception
        '    Stop
        'End Try

    End Sub

    Private Sub PerformSave_CallbackPhone(ByRef dr As DataRow, ByVal CallbackID As String, ByVal Num As Integer)
        'Try

        '    If Num = 2 Then

        '        ' ___ Work phone
        '        If Common.IsBlank(dr("WorkPhone")) Then
        '            dr("WorkPhone") = "555555555"
        '        End If

        '        PerformSave_CallbackPhone2(CallbackID, dr("WorkPhone"), "Work", 1)

        '    ElseIf Num = 1 Then

        '        ' ___ Home phone
        '        If Not Common.IsBlank(dr("HomePhone")) Then
        '            PerformSave_CallbackPhone2(CallbackID, dr("WorkPhone"), "Home", 2)
        '        End If

        '    End If

        'Catch ex As Exception
        '    Throw New Exception("Error #659: CallbackMaintain PerformSave_CallbackPhone. " & ex.Message, ex)
        'End Try
    End Sub

    Private Sub PerformSave_CallbackPhone2(ByVal CallbackID As Integer, ByVal PhoneNumber As String, ByVal Type As String, ByVal Seq As Integer)
        'Dim ChangeDate As String
        'Dim RecID As Integer
        'Dim Sql As New System.Text.StringBuilder

        'ChangeDate = Common.GetFullServerDateTime

        'RecID = Common.GetNewRecordID("CallbackPhone", "CallbackPhoneID")
        'Sql.Append("INSERT INTO CallbackPhone (CallbackPhoneID, CallbackID, PhoneNumber, PhoneExtension, Type, BestTime, Seq, LogicalDelete, AddDate, ChangeDate) ")
        'Sql.Append(" VALUES ")

        'Sql.Append("(")
        'Sql.Append(RecID.ToString & ", ")
        'Sql.Append(CallbackID & ", ")
        'Sql.Append(Common.StrOutHandler(Common.DBFormatPhone(PhoneNumber), False, StringTreatEnum.SideQts) & ", ")
        'Sql.Append("'', ")
        'Sql.Append(Common.StrOutHandler(Type, False, StringTreatEnum.SideQts) & ", ")
        'Sql.Append(Common.StrOutHandler(Type & " hours", False, StringTreatEnum.SideQts) & ", ")
        'Sql.Append(Seq.ToString & ", ")
        'Sql.Append("0, ")
        'Sql.Append(Common.DateOutHandler(ChangeDate, False, True) & ", ")
        'Sql.Append(Common.DateOutHandler(ChangeDate, False, True))
        'Sql.Append(") ")

        'Common.ExecuteNonQuery(Sql.ToString)

    End Sub

    Public Function ExcelToDT(ByVal FullPath As String, ByVal Header As Boolean, Optional ByVal Sql As String = "") As DataTable
        '   Try
        Dim dt As New DataTable

        If Sql.Length = 0 Then
            Sql = "SELECT * FROM [Sheet1$]"
        End If

        'Dim da As New OleDbDataAdapter("SELECT * FROM [Feed_DTS$]", ConnStr)
        '   Dim da As New System.Data.OleDb.OleDbDataAdapter("SELECT * FROM [ApptLicFieldsOnly$]", GetExcelConnectionString("License.xls", True))
        ' Dim da As New System.Data.OleDb.OleDbDataAdapter("SELECT * FROM [" & WorksheetName & "$]", GetExcelConnectionString(Filename, Header))
        Dim da As New System.Data.OleDb.OleDbDataAdapter(Sql, GetExcelConnectionString(FullPath, Header))
        da.Fill(dt)
        Return dt

        'Catch ex As Exception
        'Stop
        'End Try
    End Function

    Private Function GetExcelConnectionString(ByVal FullPath As String, ByVal Header As Boolean) As String
        ' http://www.connectionstrings.com/?carrier=excel

        Dim ConnString As String
        ConnString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=<fullpath>;Extended Properties=""Excel 8.0;HDR=<header>"";"
        ConnString = Replace(ConnString, "<fullpath>", FullPath)
        If Header Then
            ConnString = Replace(ConnString, "<header>", "Yes")
        Else
            ConnString = Replace(ConnString, "<header>", "No")
        End If
        Return ConnString
    End Function

    Public Sub UpdateWinEndDate()
        'Dim i As Integer
        'Dim dt As DataTable

        'Try

        '    ' dt = Common.GetDT("Select CallbackID FROM CallbackMaster where clientid = 'Diopitt' and projectreports.dbo.ufn_IsDateEqual(EnrollWinEndDate,  '12/21/2010') = 1 and LogicalDelete = 0", "hbg-sql", "Callback")
        '    dt = Common.GetDT("Select CallbackID FROM CallbackMaster where clientid = 'Diopitt' and projectreports.dbo.ufn_IsDateEqual(EnrollWinEndDate,  '12/22/2010') = 1", "hbg-sql", "Callback")
        '    For i = 0 To dt.Rows.Count - 1
        '        System.Diagnostics.Debug.WriteLine(i.ToString)
        '        Common.ExecuteNonQuery("hbg-sql", "Callback", "Update CallbackMaster Set EnrollWinEndDate = '12/23/2010' where CallbackID = " & dt.Rows(i)(0))
        '    Next

        'Catch ex As Exception
        '    Stop
        'End Try
    End Sub

    Public Sub AnotherUpdateWinEndDate()
        'Dim i As Integer
        'Dim dt As DataTable

        'Try

        '    dt = Common.GetDT("Select CallbackID FROM CallbackMaster where clientid = 'Martinrea' and projectreports.dbo.ufn_IsDateEqual(EnrollWinEndDate,  '1/3/2011') = 1 and LogicalDelete = 0", "hbg-sql", "Callback")
        '    For i = 0 To dt.Rows.Count - 1
        '        System.Diagnostics.Debug.WriteLine(i.ToString)
        '        Common.ExecuteNonQuery("hbg-sql", "Callback", "Update CallbackMaster Set EnrollWinEndDate = '1/7/2011' where CallbackID = " & dt.Rows(i)(0))
        '    Next

        'Catch ex As Exception
        '    Stop
        'End Try
    End Sub


    Public Sub UpdateStantecEmpID()
        'Dim i As Integer
        'Dim dt As DataTable
        'Dim OldEmpID As String
        'Dim NewEmpID As String

        'Try

        '    dt = Common.GetDT("Select EmpID, EmpID_Old FROM Stantec..TblEmpID_Mapping", "hbg-sql", "Stantec")
        '    For i = 0 To dt.Rows.Count - 1
        '        OldEmpID = dt.Rows(i)("EmpID_Old")
        '        NewEmpID = dt.Rows(i)("EmpID")
        '        System.Diagnostics.Debug.WriteLine(i.ToString)
        '        Common.ExecuteNonQuery("hbg-sql", "Callback", "Update CallbackMaster Set EmpID = '" & NewEmpID & "' WHERE ClientID = 'Stantec' AND EmpID = '" & OldEmpID & "' AND LogicalDelete = 0")
        '    Next

        'Catch ex As Exception
        '    Stop
        'End Try
    End Sub
End Class
