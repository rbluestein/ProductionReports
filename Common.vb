Imports System.Data.SqlClient

Module [Global]
    Private cEnviro As Enviro
    Public Property gEnviro() As Enviro
        Get
            Return cEnviro
        End Get
        Set(ByVal Value As Enviro)
            cEnviro = Value
        End Set
    End Property
End Module

Public Enum ReportNameEnum
    SupervisorReport = 1
End Enum

Public Enum BasisEnum
    Zero = 0
    One = 1
End Enum

Public Class Common
    Private cEnviro As Enviro
    Private cDiagMode As Boolean = False

    Public Enum StringTreatEnum
        AsIs = 1
        SideQts = 2
        SecApost = 3
        SideQts_SecApost = 4
    End Enum

    Public Sub New()
        cEnviro = gEnviro
    End Sub

    Public ReadOnly Property DiagMode() As Boolean
        Get
            Return cDiagMode
        End Get
    End Property

    Public Function Right(ByVal Str As String, ByVal Len As Integer) As String
        Return Str.Substring(Str.Length - Len)
    End Function

    Public Sub GenerateError()
        Dim a, b, c As Integer
        a = b / c
    End Sub

    Public Sub ExitApplication()
        Environment.Exit(0)
    End Sub

    Public Function GetReportDate(ByVal ReportConfigReportDate As String, ByVal OverrideReportDate As String) As Date
        Dim ReportDate As Date

        If OverrideReportDate = Nothing Then
            ReportDate = ReportConfigReportDate
            If ReportDate.DayOfWeek = DayOfWeek.Monday Then
                ReportDate = ReportDate.AddDays(-3)
            Else
                ReportDate = ReportDate.AddDays(-1)
            End If
        Else
            ReportDate = OverrideReportDate
        End If
        Return ReportDate
    End Function

    Public Function GetMonthName(ByVal MonthNum As Integer) As String
        Select Case MonthNum
            Case 1 : Return "Jan"
            Case 2 : Return "Feb"
            Case 3 : Return "Mar"
            Case 4 : Return "Apr"
            Case 5 : Return "May"
            Case 6 : Return "Jun"
            Case 7 : Return "Jul"
            Case 8 : Return "Aug"
            Case 9 : Return "Sep"
            Case 10 : Return "Oct"
            Case 11 : Return "Nov"
            Case 12 : Return "Dec"
        End Select
    End Function


    'Public Function GetText(ByVal Value As Object) As String
    '    Dim Results As String
    '    If IsDBNull(Value) Then
    '        Results = "[Null]"
    '    Else
    '        Results = Replace(Value, "'", "''")
    '    End If
    '    If Results = Nothing Then
    '        Results = "[EmptyString]"
    '    End If
    '    Return Results
    'End Function

#Region " Data "
    Public Sub PerformActionQuery(ByVal Sql As String)
        PerformActionQuery(Sql, cEnviro.DBHost, cEnviro.DBName)
    End Sub

    Public Sub PerformActionQuery(ByVal Sql As String, ByVal DBHost As String, ByVal DBName As String)
        Dim Querypack As QueryPack
        Querypack = PerformQueryMaster(Sql, False, False, True, False)
    End Sub

    Public Function DoesTableExist(ByVal DBName As String, ByVal TableName As String) As Boolean
        Dim sb As New System.Text.StringBuilder
        Dim dt As DataTable

        sb.Append("USE " & DBName & " ")
        sb.Append("IF EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE='BASE TABLE' AND TABLE_NAME='")
        sb.Append(TableName)
        sb.Append("') ")
        sb.Append("SELECT Results = 1 ")
        sb.Append("ELSE ")
        sb.Append("SELECT Results = 0")

        dt = GetDT(sb.ToString)
        If dt.Rows(0)(0) = 1 Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function DoesDatabaseAndTableExist(ByVal DBName As String, ByVal TableName As String) As Boolean
        Dim sb As New System.Text.StringBuilder
        Dim QueryPack As QueryPack

        sb.Append("USE " & DBName & " ")
        sb.Append("IF EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE='BASE TABLE' AND TABLE_NAME='")
        sb.Append(TableName)
        sb.Append("') ")
        sb.Append("SELECT Results = 1 ")
        sb.Append("ELSE ")
        sb.Append("SELECT Results = 0")

        QueryPack = GetDTWithQuerypack(sb.ToString)
        If QueryPack.Success Then
            If QueryPack.dt.rows(0)(0) = 1 Then
                Return True
            Else
                Return False
            End If
        Else
            Return False
        End If
    End Function

    Public Function GetFieldList(ByVal TableName As String) As DataTable
        Dim Sql As New System.Text.StringBuilder
        Sql.Append("SELECT column_name,ordinal_position,column_default,data_type, ")
        Sql.Append("Is_nullable from information_schema.columns ")
        Sql.Append("WHERE table_name='" & TableName & "'")
        Return GetDT(Sql.ToString)
    End Function

    Public Function GetDT(ByVal Sql As String) As DataTable
        Return GetDT(Sql, cEnviro.DBName, cEnviro.DBHost)
    End Function

    Public Function GetDT(ByVal Sql As String, ByVal DBName As String) As DataTable
        Return GetDT(Sql, DBName, cEnviro.DBHost)
    End Function

    Public Function GetDT(ByVal Sql As String, ByVal DBName As String, ByVal DBHost As String) As DataTable
        'Dim DataAdapter As SqlDataAdapter
        Dim dt As New DataTable
        Dim Report As Report
        Dim Querypack As QueryPack

        Try
            Querypack = PerformQueryMaster(Sql, True, False, False, False)
            Return Querypack.dt

        Catch ex As Exception
            Report = New Report
            Report.Report("Error #2202 Common.GetDT" & Sql & " " & ex.Message, Report.ReportTypeEnum.LogError)
        End Try


        'Try

        '    Dim SqlCmd As New SqlCommand(Sql)
        '    SqlCmd.CommandType = CommandType.Text
        '    SqlCmd.Connection = New SqlConnection(cEnviro.ConnectionString(DBHost, DBName))
        '    DataAdapter = New SqlDataAdapter(SqlCmd)

        '    DataAdapter.Fill(dt)

        '    DataAdapter.Dispose()
        '    SqlCmd.Dispose()
        '    Return dt

        'Catch ex As Exception
        '    Report = New Report
        '    Report.Report("Error #2202 Common.GetDT" & Sql & " " & ex.Message, Report.ReportTypeEnum.LogError)
        'End Try
    End Function

    'Public Function ORIGINALGetDTWithQueryPack(ByVal Sql As String) As QueryPack
    '    Dim DataAdapter As SqlDataAdapter
    '    Dim dt As New DataTable
    '    Dim QueryPack As New QueryPack

    '    Dim SqlCmd As New SqlCommand(Sql)
    '    SqlCmd.CommandType = CommandType.Text
    '    SqlCmd.Connection = New SqlConnection(cEnviro.ConnectionString)
    '    DataAdapter = New SqlDataAdapter(SqlCmd)
    '    Try
    '        DataAdapter.Fill(dt)
    '        QueryPack.Success = True
    '        QueryPack.dt = dt
    '    Catch ex As Exception
    '        QueryPack.Success = False
    '        QueryPack.TechErrMsg = ex.Message
    '    End Try

    '    DataAdapter.Dispose()
    '    SqlCmd.Dispose()
    '    Return QueryPack
    'End Function

    'Public Function ExecuteNonQueryWithQuerpack(ByVal Sql As String) As QueryPack
    '    Dim DataAdapter As SqlDataAdapter
    '    Dim dt As New DataTable
    '    Dim QueryPack As New QueryPack
    '    Dim Box As String()

    '    Dim SqlCmd As New SqlCommand(Sql)
    '    SqlCmd.CommandTimeout = 0
    '    SqlCmd.CommandType = CommandType.Text
    '    SqlCmd.Connection = New SqlConnection(cEnviro.ConnectionString)
    '    DataAdapter = New SqlDataAdapter(SqlCmd)
    '    Try
    '        DataAdapter.Fill(dt)
    '        Box = Split(dt.Rows(0)("RetVal"), "|")
    '        If Box(0) = 0 Then
    '            QueryPack.Success = True
    '        Else
    '            QueryPack.Success = False
    '            QueryPack.TechErrMsg = Box(1)
    '        End If

    '    Catch ex As Exception
    '        QueryPack.Success = False
    '        QueryPack.TechErrMsg = ex.Message
    '    End Try

    '    DataAdapter.Dispose()
    '    SqlCmd.Dispose()
    '    Return QueryPack
    'End Function

    'Public Function GetDTWithQueryPack(ByVal Sql As String) As QueryPack
    '    Dim DataAdapter As SqlDataAdapter
    '    Dim ds As New DataSet
    '    Dim QueryPack As New QueryPack

    '    Dim SqlCmd As New SqlCommand(Sql)
    '    SqlCmd.CommandType = CommandType.Text
    '    SqlCmd.Connection = New SqlConnection(cEnviro.ConnectionString)
    '    DataAdapter = New SqlDataAdapter(SqlCmd)


    '    Try
    '        DataAdapter.Fill(ds)
    '        If ds.Tables(0).Rows(0)("ErrorNumber") = 0 Then
    '            QueryPack.Success = True
    '            QueryPack.dt = ds.Tables(1)
    '        Else
    '            QueryPack.Success = False
    '            QueryPack.TechErrMsg = ds.Tables(0).Rows(0)("ErrorMessage")
    '        End If

    '    Catch ex As Exception
    '        QueryPack.Success = False
    '        QueryPack.TechErrMsg = ex.Message
    '    End Try

    '    DataAdapter.Dispose()
    '    SqlCmd.Dispose()
    '    Return QueryPack
    'End Function

    'Public Function GetDSWithQueryPack(ByVal Sql As String) As QueryPack
    '    Dim i As Integer
    '    Dim DataAdapter As SqlDataAdapter
    '    Dim ds As New DataSet
    '    Dim QueryPack As New QueryPack

    '    Dim SqlCmd As New SqlCommand(Sql)
    '    SqlCmd.CommandType = CommandType.Text
    '    SqlCmd.Connection = New SqlConnection(cEnviro.ConnectionString)
    '    DataAdapter = New SqlDataAdapter(SqlCmd)


    '    Try
    '        DataAdapter.Fill(ds)

    '        If ds.Tables(0).Rows(0)("ErrorNumber") = 0 Then
    '            QueryPack.Success = True
    '            ds.Tables.Remove(ds.Tables(0))
    '            QueryPack.ds = ds
    '        Else
    '            QueryPack.Success = False
    '            QueryPack.TechErrMsg = ds.Tables(0).Rows(0)("ErrorMessage")
    '        End If

    '    Catch ex As Exception
    '        QueryPack.Success = False
    '        QueryPack.TechErrMsg = ex.Message
    '    End Try

    '    DataAdapter.Dispose()
    '    SqlCmd.Dispose()
    '    Return QueryPack
    'End Function

    Public Function GetDTWithQuerypack(ByVal Sql As String) As QueryPack
        Return PerformQueryMaster(Sql, True, False, False, False)
    End Function

    Public Function GetDSWithQuerypack(ByVal Sql As String) As QueryPack
        Return PerformQueryMaster(Sql, False, True, False, False)
    End Function

    Public Function ExecuteNonQueryWithQuerypack(ByVal Sql As String) As QueryPack
        Return PerformQueryMaster(Sql, False, False, True, False)
    End Function

    Public Function ExecuteNonQuerySPErrorHandlingWithQuerypack(ByVal Sql As String) As QueryPack
        Return PerformQueryMaster(Sql, False, False, True, True)
    End Function

    Public Function TestServerConnection(ByVal Timeout As Integer) As Boolean
        Dim Querypack As QueryPack
        Querypack = PerformQueryMaster("Select Count (*) FROM EmpTransmittal", False, False, True, False, Timeout)
        If Querypack.Success Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function PerformQueryMaster(ByVal Sql As String, ByVal WithDatatable As Boolean, ByVal WithDataset As Boolean, ByVal ActionQuery As Boolean, ByVal WithSPErrorHandling As Boolean, Optional ByVal CommandTimeout As Integer = 600) As QueryPack
        Dim DataAdapter As SqlDataAdapter
        Dim dt As New DataTable
        Dim ds As New DataSet
        Dim QueryPack As New QueryPack
        Dim Box As String()
        Dim SqlCmd As SqlCommand
        Dim NumResultsExpected As String

        Try

            SqlCmd = New SqlCommand(Sql)
            SqlCmd.CommandTimeout = CommandTimeout
            SqlCmd.CommandType = CommandType.Text
            SqlCmd.Connection = New SqlConnection(cEnviro.ConnectionString)
            DataAdapter = New SqlDataAdapter(SqlCmd)

            ' ___ Are we expecting results?
            If ActionQuery AndAlso (Not WithSPErrorHandling) Then
                NumResultsExpected = "none"
            ElseIf WithDatatable And (Not WithSPErrorHandling) Then
                NumResultsExpected = "one"
            Else
                NumResultsExpected = "mult"
            End If

            ' ___ Perform the query
            Select Case NumResultsExpected
                Case "none", "one"
                    'Try
                    DataAdapter.Fill(dt)
                    'Catch
                    'End Try
                Case "mult"
                    DataAdapter.Fill(ds)
            End Select

            ' ___ If error handling expected and not returned, return error.
            If WithSPErrorHandling Then
                Select Case NumResultsExpected
                    Case "one"
                        If dt Is Nothing Then
                            QueryPack.Success = False
                            QueryPack.TechErrMsg = "Stored procedure error message not returned."
                            Return QueryPack
                        End If
                    Case "mult"
                        If ds.Tables.Count = 0 Then
                            QueryPack.Success = False
                            QueryPack.TechErrMsg = "Stored procedure error message not returned."
                            Return QueryPack
                        End If
                End Select
            End If

            ' ___ Process the stored procedure error handling.
            If WithSPErrorHandling Then
                Box = Split(ds.Tables(0).Rows(0)(0), "|")
                If Box(0) = 0 Then
                    QueryPack.Success = True
                    ds.Tables.Remove(ds.Tables(0))
                Else
                    QueryPack.Success = False
                    QueryPack.TechErrMsg = Box(1)
                    Return QueryPack
                End If
            End If
            ' //  Package and return the resultsets.


            If WithDatatable AndAlso WithSPErrorHandling Then

                ' ___ If query submitted as WithDataTable as well as WithSPErrorHandling, return datatable, not dataset.
                If ds.Tables.Count > 0 Then
                    QueryPack.dt = ds.Tables(0)
                End If

            ElseIf WithDataset Then
                QueryPack.ds = ds
            ElseIf WithDatatable Then
                QueryPack.dt = dt
            End If

            QueryPack.Success = True
            Return QueryPack


        Catch ex As Exception
            QueryPack.Success = False
            QueryPack.TechErrMsg = ex.Message
            Return QueryPack
        Finally
            DataAdapter.Dispose()
            SqlCmd.Dispose()
        End Try
    End Function


    'Sub DisplayUserSelection(ByVal OcsIdList As String)
    '    Dim i As Integer
    '    Dim RegSessions As Integer
    '    Dim Docs As Integer
    '    Dim Policies As Integer
    '    Dim OcsIdArray As String()
    '    Dim HdrColl As New Collection
    '    Dim DtlColl As New Collection
    '    Dim szb As New System.Text.StringBuilder
    '    Dim Selection As Integer

    '    Selection = Request.Form("ddSelect")

    '    ' ----------------------------------------------------------
    '    ' Each row in the header collection is a table. Each table 
    '    ' has 1 row corresponding to a reg session's header data.                                                           
    '    ' Each row in the detail collection is a table. For reg sessions, 
    '    ' each table has 1 row corresponding to a reg session's field data.
    '    ' For docs and policies, each resultset has a list of the
    '    ' docs or policies for that registration session.
    '    ' ----------------------------------------------------------

    '    SqlCmdRead.Parameters("@ClubId").Value = CInt(oSession.fnGetSession("OrgId"))
    '    OcsIdArray = Split(OcsIdList, "|")
    '    SqlCmdRead.Parameters("@Action").Value = Selection

    '    For i = 0 To OcsIdArray.GetUpperBound(0)
    '        If Not Request("OcsId" & CStr(OcsIdArray(i))) = Nothing Then

    '            SqlCmdRead.Parameters("@OcsId").Value = OcsIdArray(i)
    '            Dim objDataAdapter As New SqlDataAdapter(SqlCmdRead)
    '            Dim ds As New DataSet
    '            objDataAdapter.Fill(ds)
    '            HdrColl.Add(ds.Tables.Item(0))
    '            DtlColl.Add(ds.Tables.Item(1))
    '            ds.Dispose()
    '            objDataAdapter.Dispose()

    '        End If
    '    Next

    '    If HdrColl.Count = 0 Then
    '        litData.Text = String.Empty
    '    Else
    '        Select Case Selection
    '            Case 2
    '                cRptType = RptType.RegSessions
    '                ProcessRegSessions(HdrColl, DtlColl, szb)
    '                litData.Text = szb.ToString
    '            Case 3
    '                cRptType = RptType.Docs
    '                ProcessDocsPolicies("Docs", HdrColl, DtlColl, szb)
    '                litData.Text = szb.ToString
    '            Case 4
    '                cRptType = RptType.Policies
    '                ProcessDocsPolicies("Policies", HdrColl, DtlColl, szb)
    '                litData.Text = szb.ToString
    '        End Select
    '    End If

    '    '        'If HdrColl.Count = 0 Then
    '    '        '    litMsg.Text = "<script Language=""JavaScript"">" & vbCrLf & _
    '    '        ' "alert('No registration sessions selected.')" & vbCrLf & _
    '    '        ' "</script>" & vbCrLf
    '    '        '    Exit Sub
    '    '        'End If


    '    SqlCmdRead.Dispose()
    '    SqlCmdLookup.Dispose()
    'End Sub


#End Region

#Region " In handlers "
    Public Function StrInHandler(ByVal Input As Object) As Object
        Dim Output As Object

        If IsDBNull(Input) Then
            Return String.Empty
        ElseIf (Not IsNumeric(Input)) AndAlso Input = Nothing Then
            Return String.Empty
            'ElseIf (Not IsDate(Input)) AndAlso Input.length = 0 Then
            '    Return String.Empty
        Else
            Output = Replace(Input, "~", "'")
            If Output = Nothing Then
                Return String.Empty
            End If
            Return Output
        End If
    End Function
    Public Function DateInHandler(ByVal Input As Object) As Object
        ' 12/31/2399
        Dim Output As Object
        Output = Input

        If IsDBNull(Input) Then
            Return String.Empty
        ElseIf Input = "01/01/1900" Then
            Return String.Empty
        ElseIf Input = "01/01/1950" Then
            Return String.Empty
        Else
            Return Output
        End If
    End Function

    Public Function NumInHandler(ByVal Input As Object, ByVal NullAsZero As Boolean) As Object
        If IsDBNull(Input) Then
            If NullAsZero Then
                Return 0
            Else
                Return String.Empty
            End If
        Else
            Return Input
        End If
    End Function

    Public Function GuidInHandler(ByVal Input As Object) As Object
        If IsDBNull(Input) Then
            Return String.Empty
        Else
            Return Input.ToString
        End If
    End Function


    'Public Function StrXferHandler(ByVal Input As Object, ByVal AllowNull As Boolean) As Object
    '    Dim Output As Object
    '    Dim ReturnNull As Boolean

    '    If IsDBNull(Input) Then
    '        ReturnNull = True
    '    ElseIf (Not IsNumeric(Input)) AndAlso Input = Nothing Then
    '        ReturnNull = True
    '    Else
    '        Output = Replace(Input, "~", "'")
    '        If Output = Nothing Then
    '            ReturnNull = True
    '        End If
    '    End If

    '    If ReturnNull Then
    '        If AllowNull Then
    '            Return DBNull.Value
    '        Else
    '            Return String.Empty
    '        End If
    '    Else
    '        Return Output
    '    End If

    'End Function
    'Public Function DateXferHandler(ByVal Input As Object, ByVal AllowNull As Boolean) As Object
    '    ' 12/31/2399
    '    Dim Output As Object
    '    Dim ReturnNull As Boolean
    '    Output = Input

    '    If IsDBNull(Input) OrElse Input = Nothing OrElse Input = "1/1/1950" Then
    '        ReturnNull = True
    '    Else
    '        Output = Input
    '    End If

    '    If ReturnNull Then
    '        If AllowNull Then
    '            Return DBNull.Value
    '        Else
    '            Return "1/1/1950"
    '        End If
    '    Else
    '        Return Output
    '    End If
    'End Function

    'Public Function NumXferHandler(ByVal Input As Object, ByVal AllowNull As Boolean) As Object
    '    Dim Output As Object
    '    Dim ReturnNull As Boolean

    '    If IsDBNull(Input) Then
    '        ReturnNull = True
    '    Else
    '        Output = Input
    '    End If

    '    If ReturnNull Then
    '        If AllowNull Then
    '            Return DBNull.Value
    '        Else
    '            Return 0
    '        End If
    '    Else
    '        Return Output
    '    End If

    'End Function
#End Region

#Region " Out handlers"
    Public Function StrOutHandler(ByRef Input As Object, ByVal AllowNull As Boolean, ByVal StringTreat As StringTreatEnum) As Object
        Dim Output As String

        Try

            ' ___ Output, adjusting for AllowNull
            If IsDBNull(Input) Then
                If AllowNull Then
                    Output = "null"
                Else
                    Output = String.Empty
                End If
            ElseIf Input Is Nothing Then
                If AllowNull Then
                    Output = "null"
                Else
                    Output = String.Empty
                End If
            Else
                Try
                    Output = Input
                Catch
                    If AllowNull Then
                        Output = "null"
                    Else
                        Output = String.Empty
                    End If
                End Try
            End If

            ' ___ Apply string treatment
            If Output <> "null" Then
                Select Case StringTreat
                    Case StringTreatEnum.AsIs
                        ' no action
                    Case StringTreatEnum.SecApost
                        Output = Replace(Output, "'", "''")
                    Case StringTreatEnum.SideQts
                        Output = "'" & Output & "'"
                    Case StringTreatEnum.SideQts_SecApost
                        Output = Replace(Output, "'", "''")
                        Output = "'" & Output & "'"
                End Select
            End If

            Return Output

        Catch ex As Exception
            Throw New Exception("Error #CM2220: Common StrOutHandler. " & ex.Message, ex)
        End Try
    End Function

    Public Function DateOutHandler(ByVal Input As Object, ByVal AllowNull As Boolean, Optional ByVal AddSingleQuotes As Boolean = False) As Object
        Dim ReturnNull As Boolean
        Dim Output As Object

        If IsDBNull(Input) OrElse Input = Nothing Then
            If AllowNull Then
                ReturnNull = True
            Else
                Output = "01/01/1950"
            End If
        Else
            Output = Input
        End If

        If ReturnNull Then
            Return "null"
        Else
            If AddSingleQuotes Then
                Return "'" & Output & "'"
            Else
                Return Output
            End If
        End If
    End Function
    Public Function PhoneOutHandler(ByVal Input As Object, ByVal AllowNull As Boolean, Optional ByVal AddSingleQuotes As Boolean = False) As Object
        Dim i As Integer
        Dim Output As String = String.Empty
        Dim Working As String
        'Working = StrOutHandler(Input, AllowNull, AddSingleQuotes)

        If AddSingleQuotes Then
            Working = StrOutHandler(Input, AllowNull, StringTreatEnum.SideQts)
        Else
            Working = StrOutHandler(Input, AllowNull, StringTreatEnum.AsIs)
        End If


        If Working = "null" Or Working = String.Empty Then
        Else
            If Working.Length >= 10 Then
                For i = 0 To Working.Length - 1
                    If IsNumeric(Working.Substring(i, 1)) Then
                        Output &= Working.Substring(i, 1)
                    End If
                Next
            End If
        End If

        If Output.Length = 10 Then
            Output = InsertAt(Output, "(", 1)
            Output = InsertAt(Output, ") ", 5)
            Output = InsertAt(Output, "-", 10)
        Else
            Output = Input
        End If

        If AddSingleQuotes Then
            Output = "'" & Output & "'"
        End If

        Return Output
    End Function

    Public Function BitOutHandler(ByVal Input As Object, ByVal AllowNull As Boolean) As Object
        If IsDBNull(Input) Then
            If AllowNull Then
                Return "null"
            Else
                Return 0
            End If
        Else
            If CType(Input, Boolean) Then
                Return 1
            Else
                Return 0
            End If
        End If
    End Function
#End Region

#Region " Validate "
    Public Function IsBlank(ByVal Value As Object) As Boolean
        If IsDBNull(Value) Then
            Return True
        ElseIf Value = Nothing Then
            Return True
        Else
            If Value.length = 0 Then
                Return True
            Else
                Return False
            End If
        End If
    End Function

    Public Sub ValidateStringField(ByRef ErrColl As Collection, ByVal Value As Object, ByVal MinLength As Integer, ByVal ErrMsg As String)
        If Value.length < MinLength Then
            If ErrColl.Count = 0 Then
                ErrColl.Add(ErrMsg)
            Else
                ErrColl.Add(", " & ErrMsg)
            End If
        End If
    End Sub

    Public Sub ValidateStringField(ByRef ErrColl As Collection, ByVal Value As Object, ByVal MinLength As Integer, ByVal MaxLength As Integer, ByVal ErrMsg As String)
        If Value.length < MinLength Or Value.length > MaxLength Then
            If ErrColl.Count = 0 Then
                ErrColl.Add(ErrMsg)
            Else
                ErrColl.Add(", " & ErrMsg)
            End If
        End If
    End Sub

    Public Sub ValidateNumericField(ByRef ErrColl As Collection, ByVal Value As Object, ByVal AllowNull As Boolean, ByVal ErrMsg As String)
        Dim PassTest As Boolean
        If IsDBNull(Value) OrElse Value.Length = 0 Then
            If AllowNull Then
                PassTest = True
            Else
                PassTest = False
            End If
        Else
            If IsNumeric(Value) Then
                PassTest = True
            Else
                PassTest = False
            End If
        End If

        If Not PassTest Then
            If ErrColl.Count = 0 Then
                ErrColl.Add(ErrMsg)
            Else
                ErrColl.Add(", " & ErrMsg)
            End If
        End If
    End Sub

    Public Sub ValidateRadio(ByRef ErrColl As Collection, ByVal SelectedIndex As Integer, ByVal AllowNull As Boolean, ByVal ErrMsg As String)
        If (SelectedIndex < 0) AndAlso (Not AllowNull) Then
            If ErrColl.Count = 0 Then
                ErrColl.Add(ErrMsg)
            Else
                ErrColl.Add(", " & ErrMsg)
            End If
        End If
    End Sub

    Public Sub ValidateNumericRange(ByRef ErrColl As Collection, ByVal Value As Object, ByVal Min As Integer, ByVal Max As Integer, ByVal AllowNull As Boolean, ByVal ErrMsg As String)
        Dim PassTest As Boolean
        If IsDBNull(Value) OrElse Value.Length = 0 Then
            If AllowNull Then
                PassTest = True
            Else
                PassTest = False
            End If
        Else
            If IsNumeric(Value) Then
                If Value >= Min AndAlso Value <= Max Then
                    PassTest = True
                Else
                    PassTest = False
                End If
            End If
        End If

        If Not PassTest Then
            If ErrColl.Count = 0 Then
                ErrColl.Add(ErrMsg)
            Else
                ErrColl.Add(", " & ErrMsg)
            End If
        End If
    End Sub

    Public Sub ValidateDateField(ByRef ErrColl As Collection, ByVal Value As Object, ByVal AllowNull As Boolean, ByVal ErrMsg As String)
        Dim Valid As Boolean
        If IsDBNull(Value) OrElse Value = Nothing Then
            If AllowNull Then
                Valid = True
            Else
                Valid = False
            End If
        ElseIf IsDate(Value) Then
            Valid = True
        Else
            Valid = False
        End If
        If Not Valid Then
            If ErrColl.Count = 0 Then
                ErrColl.Add(ErrMsg)
            Else
                ErrColl.Add(", " & ErrMsg)
            End If
        End If
    End Sub

    Public Function IsValidPhoneNumber(ByVal Value As Object) As Boolean
        Dim i As Integer
        Dim NumCount As Integer

        If IsDBNull(Value) OrElse Value = Nothing Then
            Return False
        End If

        If Value.length >= 10 Then
            For i = 0 To Value.Length - 1
                If IsNumeric(Value.Substring(i, 1)) Then
                    NumCount += 1
                End If
            Next
        End If

        If NumCount = 10 Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Sub ValidatePhoneNumber(ByRef Errcoll As Collection, ByVal Value As Object, ByVal ErrMsg As String)
        If Not IsValidPhoneNumber(Value) Then
            If Errcoll.Count = 0 Then
                Errcoll.Add(ErrMsg)
            Else
                Errcoll.Add(", " & ErrMsg)
            End If
        End If
    End Sub

    Public Sub ValidateEmailAddress(ByRef ErrColl As Collection, ByVal Value As Object, ByVal ErrMsg As String)
        If Not IsValidEmailAddress(Value) Then
            If ErrColl.Count = 0 Then
                ErrColl.Add(ErrMsg)
            Else
                ErrColl.Add(", " & ErrMsg)
            End If
        End If
    End Sub

    Public Function IsValidEmailAddress(ByVal Value As Object) As Boolean
        Dim OKSoFar As Boolean = True
        Const InvalidChars As String = "!#$%^&*()=+{}[]|\;:'/?>,< "
        Dim i As Integer
        Dim Num As Integer
        Dim DotPos As Integer
        Dim Part2 As String

        ' ___ Check for null or empty value
        If IsDBNull(Value) OrElse Value = Nothing Then
            OKSoFar = False
        End If

        ' ___ Check for minimum length
        If OKSoFar Then
            If Value.Length < 5 Then
                OKSoFar = False
            End If
        End If

        ' ___ Check for a double quote
        If OKSoFar Then
            OKSoFar = Not InStr(1, Value, Chr(34)) > 0  'Check to see if there is a double quote
        End If

        ' ___ Check for consecutive dots
        If OKSoFar Then
            OKSoFar = Not InStr(1, Value, "..") > 0
        End If

        ' ___ Check for invalid characters
        If OKSoFar Then
            For i = 0 To InvalidChars.Length - 1
                If InStr(1, Value, InvalidChars.Substring(i, 1)) > 0 Then
                    OKSoFar = False
                    Exit For
                End If
            Next
        End If

        ' ___ Check for number of @ symbols
        If OKSoFar Then
            For i = 0 To Value.Length - 1
                If InStr(Value.Substring(i, 1), "@") > 0 Then
                    Num += 1
                End If
            Next
            If Num > 1 Then
                OKSoFar = False
            End If
        End If

        ' ___ Check for the @ symbol in starting before the third position
        If OKSoFar Then
            If InStr(Value, "@") < 2 Then
                OKSoFar = False
            End If
        End If

        ' ___ Check for number of dots
        If OKSoFar Then
            Num = 0
            Part2 = Value.substring(InStr(Value, "@"))
            For i = 0 To Part2.Length - 1
                If InStr(Part2.Substring(i, 1), ".") > 0 Then
                    Num += 1
                End If
            Next
            If Num > 1 Then
                OKSoFar = False
            End If
        End If

        ' ___ Dot is present and not immediately after ampersand and not at end. 
        '___  Dot separated from ampersand by at least one character
        If OKSoFar Then
            DotPos = InStr(Part2, ".")
            If DotPos < 2 Or DotPos = Part2.Length Then
                OKSoFar = False
            End If
        End If

        Return OKSoFar
    End Function


    Public Sub ValidateCheckbox(ByRef ErrColl As Collection, ByRef chkBox As CheckBox, ByVal ValidState As Integer, ByVal ErrMsg As String)
        Dim IsValid As Boolean = True
        If ValidState = 0 AndAlso Not chkBox.Checked Then
            IsValid = False
        ElseIf ValidState = 1 AndAlso Not chkBox.Checked Then
            IsValid = False
        End If
        If Not IsValid Then
            If ErrColl.Count = 0 Then
                ErrColl.Add(ErrMsg)
            Else
                ErrColl.Add(", " & ErrMsg)
            End If
        End If
    End Sub

    Public Sub ValidateErrorOnly(ByRef ErrColl As Collection, ByVal ErrMsg As String)
        If ErrColl.Count = 0 Then
            ErrColl.Add(ErrMsg)
        Else
            ErrColl.Add(", " & ErrMsg)
        End If
    End Sub

    Public Function ErrCollToString(ByRef ErrColl As Collection, ByVal Intro As String) As String
        Dim sb As New System.Text.StringBuilder
        Dim i As Integer
        If ErrColl.Count > 0 Then
            For i = 1 To ErrColl.Count
                sb.Append(ErrColl(i))
            Next
        End If
        Return Intro & " " & sb.ToString & "."
    End Function
#End Region

#Region " This to that "
    Public Function BitToRadio(ByVal Value As Object, ByVal TrueIndex As Integer, ByVal AllowNoneSelected As Boolean) As Integer
        Dim FalseIndex As Integer
        FalseIndex = System.Math.Abs(TrueIndex - 1)

        If IsDBNull(Value) Then
            If AllowNoneSelected Then
                Return -1
            Else
                Return FalseIndex
            End If
        Else
            If Value Then
                Return TrueIndex
            Else
                Return FalseIndex
            End If
        End If
    End Function

    Public Function BitToString(ByVal Value As Object, ByVal TrueString As String, ByVal FalseString As String, ByVal AllowNull As Boolean) As String
        If IsDBNull(Value) Then
            If AllowNull Then
                Return String.Empty
            Else
                Return FalseString
            End If
        End If
        If Value Then
            Return TrueString
        Else
            Return FalseString
        End If
    End Function

    Public Function ChkToInd(ByVal chkBox As CheckBox) As Integer
        If chkBox.Checked Then
            Return 1
        Else
            Return 0
        End If
    End Function
    Public Sub IndToChk(ByVal Ind As Object, ByVal chkBox As CheckBox)
        If IsDBNull(Ind) Then
            chkBox.Checked = False
        Else
            If Ind Then
                chkBox.Checked = True
            Else
                chkBox.Checked = False
            End If
        End If
    End Sub
#End Region

#Region " Everything else "

    'Public Sub RecordLoggedInUserData(ByVal LoggedInUserID As String, ByVal SessionID As String, ByVal LastLoginIP As String)
    '    Dim dt As DataTable
    '    Dim Sql As String

    '    dt = GetDT("SELECT LastSessionID FROM Users WHERE UserID = '" & LoggedInUserID & "'")
    '    If dt.Rows(0)("LastSessionID") <> SessionID Then
    '        Sql = "UPDATE Users SET LastSessionID = '" & SessionID & "', LastLoginDate = '" & GetServerDateTime() & "', LastLoginIP = '" & LastLoginIP & "' WHERE UserID = '" & LoggedInUserID & "'"
    '        ExecuteNonQuery(Sql)
    '    End If
    'End Sub

    Public Function RowItemFind(ByRef dt As DataTable, ByVal ColumnNum As Integer, ByVal Item As Object) As Integer
        Dim i As Integer
        Dim Results As Integer = -1

        For i = 0 To dt.Rows.Count - 1
            If dt.Rows(i)(ColumnNum) = Item Then
                Return i
            End If
        Next
        Return Results
    End Function

    Public Function NumPart(ByVal Value As String) As Integer
        Dim i As Integer
        Dim Result As String
        For i = 0 To Value.Length - 1
            If Asc(Value.Substring(i, 1)) > 47 And Asc(Value.Substring(i, 1)) < 58 Then
                Result &= Value.Substring(i, 1)
            End If
        Next
        If Result = Nothing Then
            Return -1
        Else
            Return CType(Result, System.Int64)
        End If
    End Function


    Public Function GetServerDateTime() As DateTime
        Return Date.Now.ToUniversalTime.AddHours(-5)
    End Function

    Public Function ConditionStringForHTML(ByVal Value As Object) As String
        Dim Results As String
        If IsDBNull(Value) Then
            Results = String.Empty
        Else
            Results = Value.ToString
        End If
        Results = Replace(Results, Chr(10).ToString, "<br />")
        Return Results
    End Function

    Public Function GetRightsStr(ByRef dt As DataTable) As String
        Dim i As Integer
        Dim sb As New System.Text.StringBuilder

        If dt.Rows.Count = 0 Then
            Return String.Empty
        Else
            For i = 0 To dt.Rows.Count - 1
                sb.Append("|" & StrInHandler(dt.Rows(i)("RightCd")))
            Next
            Return sb.ToString.Substring(1)
        End If
    End Function

    Public Function InsertAt(ByVal Value As String, ByVal InsChar As String, ByVal Pos As Integer) As String
        Dim ValuePos As Integer = 1
        Dim Output As String = String.Empty
        Dim OutputPos As Integer = 1
        Do
            If OutputPos = Pos Then
                Output &= InsChar
                OutputPos += 1
            Else
                Output &= Value.Substring(ValuePos - 1, 1)
                ValuePos += 1
                OutputPos += 1
                If ValuePos > Value.Length Then
                    Exit Do
                End If
            End If
        Loop
        Return Output
    End Function

    Public Function ToNull(ByVal Input As Object) As Object
        If IsDBNull(Input) Then
            Return DBNull.Value
        ElseIf Input Is Nothing Then
            Return DBNull.Value
        ElseIf Input.length = 0 Then
            Return DBNull.Value
        Else
            Return Input
        End If
    End Function



    'Public Function DateSqlWhere(ByRef Input As Object) As String
    '    If IsDBNull(Input) Then
    '        Return " is null "
    '    ElseIf Input = Nothing Then
    '        Return " is null "
    '    ElseIf Trim(Input).Length = 0 Then
    '        Return " is null "
    '    Else
    '        Return " = '" & Input & "' "
    '    End If
    'End Function

    'Public Function DateSqlWhereNoNull(ByRef Input As Object) As String
    '    If IsDBNull(Input) Then
    '        Return " = '01/01/1900' "
    '    ElseIf Input = Nothing Then
    '        Return " = '01/01/1900' "
    '    ElseIf Trim(Input).Length = 0 Then
    '        Return " = '01/01/1900' "
    '    Else
    '        Return " = '" & Input & "' "
    '    End If
    'End Function

    Public Function IsBVIDate(ByVal Input As Object) As Boolean
        If IsDBNull(Input) Then
            Return False
        ElseIf Input = Nothing Then
            Return False
        ElseIf Input = "01/01/1950" Then
            Return False
        ElseIf Input.ToString = String.Empty Then
            Return False
        Else
            Return True
        End If
    End Function


    Public Function IsDateBetween(ByVal FromDate As Object, ByVal ToDate As Object, ByVal SubjectDate As DateTime) As Boolean
        Dim AfterStart As Boolean
        Dim BeforeEnd As Boolean
        Dim Results As Boolean

        If Not IsDate(FromDate) Then
            AfterStart = True
        Else
            If DateCompare(SubjectDate, FromDate, False) > -1 Then
                AfterStart = True
            End If
        End If

        If Not IsDate(ToDate) Then
            BeforeEnd = True
        Else
            If DateCompare(SubjectDate, ToDate, False) < 1 Then
                BeforeEnd = True
            End If
        End If

        If AfterStart And BeforeEnd Then
            Results = True
        End If

        Return Results
    End Function

    Public Function DateCompare(ByVal Date1 As DateTime, ByVal Date2 As DateTime, ByVal CompareDatePartOnly As Boolean) As Single
        Try
            If CompareDatePartOnly Then
                Date1 = CType(Date1.ToString("MM/dd/yyyy") & " 00:00 AM", System.DateTime)
                Date2 = CType(Date2.ToString("MM/dd/yyyy") & " 00:00 AM", System.DateTime)
            End If
            Return Date.Compare(Date1, Date2)
        Catch ex As Exception
            Throw New Exception("Error #CM2210: Common DateCompare " & ex.Message, ex)
        End Try
    End Function

    Public Function GetFirstBusinessDateOfYear(ByVal Year As Integer) As DateTime
        Dim FirstBusinessDate As DateTime
        Dim DayOfWeek As Integer

        FirstBusinessDate = CType("1/2/" & CType(Year, System.String), System.DateTime)
        DayOfWeek = FirstBusinessDate.DayOfWeek
        Select Case DayOfWeek
            Case 0  ' Sunday to Monday
                FirstBusinessDate = FirstBusinessDate.AddDays(1)
            Case 6 ' Saturday to Monday
                FirstBusinessDate = FirstBusinessDate.AddDays(2)
        End Select

        Return FirstBusinessDate
    End Function

    Public Function GetNewRecordID(ByVal TableName As String, ByVal KeyFldName As String) As Results
        Dim MyResults As New Results
        Dim RandValue As Integer
        Dim dt As DataTable

        Try

            Do
                Do
                    Randomize()
                    RandValue = CType(Rnd() * 1000000, System.Int64)
                Loop Until RandValue > 99999
                dt = GetDT("SELECT Count (*) FROM " & TableName & " WHERE " & KeyFldName & " = " & RandValue)
                If dt.Rows(0)(0) = 0 Then
                    Exit Do
                End If
            Loop

            MyResults.Success = True
            MyResults.Value = RandValue
            Return MyResults

            'Return RandValue

        Catch ex As Exception
            'Report = New Report
            'Report.Report("Main #100 " & ex.Message, Report.ReportTypeEnum.LogError)

            MyResults.Success = False
            MyResults.Message = "Common.GetNewRecordID " & ex.Message
            Return MyResults
        End Try

    End Function

    Public Function GetRandomlyGeneratedPassword(ByVal Length As Integer) As String
        Dim i As Integer
        Dim RndValue As Integer
        Dim sb As New System.Text.StringBuilder

        ' ___ Generate random password 8 digits long
        For i = 1 To Length
            Randomize()
            RndValue = CInt(Int(62 * Rnd() + 1))
            Select Case RndValue
                Case 1 To 10
                    sb.Append(Chr(RndValue + 47))
                Case 11 To 36
                    sb.Append(Chr(RndValue + 54))
                Case 37 To 62
                    sb.Append(Chr(RndValue + 28))
            End Select
        Next
        Return sb.ToString.ToLower

    End Function

    Public Function GetCurRightsHidden(ByVal RightsColl As Collection) As String
        Dim i As Integer
        Dim sb As New System.Text.StringBuilder
        For i = 1 To RightsColl.Count
            sb.Append(RightsColl(i) & "|")
        Next
        sb.Length -= 1
        Return "<input type='hidden' id='currentrights' name='currentrights' value=""" & sb.ToString & """>"
    End Function

    Public Function GetCurRightsHidden(ByVal RightsStr As String) As String
        Return "<input type='hidden'id='currentrights' name='currentrights' value=""" & RightsStr & """ > "
    End Function
#End Region

End Class
