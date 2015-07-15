Imports System.IO
Imports System.Data.SqlClient

Public Class Results
    Private cSuccess As Boolean
    Private cMessage As String
    Private cValue As Object

    Public Property Success() As Boolean
        Get
            Return cSuccess
        End Get
        Set(ByVal Value As Boolean)
            cSuccess = Value
        End Set
    End Property
    Public Property Message() As String
        Get
            Return cMessage
        End Get
        Set(ByVal Value As String)
            cMessage = Value
        End Set
    End Property
    Public Property Value() As Object
        Get
            Return cValue
        End Get
        Set(ByVal Value As Object)
            cValue = Value
        End Set
    End Property
End Class

Public Class QueryPack
    Private cReturnDataTable As Boolean
    Private cReturnDataSet As Boolean
    Private cSuccess As Boolean
    Private cGenErrMsg As String
    Private cTechErrMsg As String
    Private cdt As DataTable
    Private cds As DataSet

    Public Property Success() As Boolean
        Get
            Return cSuccess
        End Get
        Set(ByVal Value As Boolean)
            cSuccess = Value
        End Set
    End Property

    Public ReadOnly Property GenErrMsg() As String
        Get
            Return GenErrMsg
        End Get
    End Property
    Public Property TechErrMsg() As String
        Get
            Return cTechErrMsg
        End Get
        Set(ByVal Value As String)
            cTechErrMsg = Value
        End Set
    End Property
    Public Property dt() As DataTable
        Get
            Return cdt
        End Get
        Set(ByVal Value As DataTable)
            cdt = Value
        End Set
    End Property
    Public Property ds() As DataSet
        Get
            Return cds
        End Get
        Set(ByVal Value As DataSet)
            cds = Value
        End Set
    End Property
End Class

Public Class CmdPack
    Private cSqlCmd As SqlClient.SqlCommand
    Private cParameterColl As New Collection
    Private cTechErrMsg As String
    Private cSuccess As Boolean

    Public Sub New(ByVal SPNameOrSqlText As String, ByVal CmdType As CommandType, ByVal Enviro As Enviro)
        cSqlCmd = New SqlClient.SqlCommand(SPNameOrSqlText)
        cSqlCmd.CommandType = CmdType
        cSqlCmd.Connection = New SqlConnection(Enviro.ConnectionString)
    End Sub

    Public Sub AddParameter(ByVal ParameterType As SqlDbType, ByVal VarName As String, ByVal Value As Object, ByVal Direction As ParameterDirection, Optional ByVal Length As Integer = 0)
        Dim Parameter As SqlParameter
        Parameter = New SqlParameter(VarName, SqlDbType.VarChar, Length)
        Parameter.Value = Value
        Parameter.Direction = Direction
        cSqlCmd.Parameters.Add(Parameter)
        If ParameterType = SqlDbType.VarChar Then
            Parameter.Size = Length
        End If
    End Sub

    Public Sub Execute()
        Dim DataReader As SqlDataReader

        Try
            cSqlCmd.Connection.Open()
            DataReader = cSqlCmd.ExecuteReader()
            Dim Param As SqlParameter
            For Each Param In cSqlCmd.Parameters
                cParameterColl.Add(Param.Value, Param.ParameterName)
            Next
            cSuccess = True

        Catch ex As Exception
            cSuccess = False
            cTechErrMsg = ex.Message
        Finally
            Try
                DataReader.Close()
            Catch
            End Try
            cSqlCmd.Dispose()
            cSqlCmd.Connection.Close()
            End Try
    End Sub

    Public ReadOnly Property Success() As Boolean
        Get
            Return cSuccess
        End Get
    End Property
    Public ReadOnly Property TechErrMsg() As String
        Get
            Return cTechErrMsg
        End Get
    End Property
    Public ReadOnly Property ParameterColl() As Collection
        Get
            Return cParameterColl
        End Get
    End Property
End Class

Public Class ExcelPack_Generic
    Dim cdt As DataTable
    Dim cRangeName As String

    Public Property DT() As DataTable
        Get
            Return cdt
        End Get
        Set(ByVal Value As DataTable)
            cdt = Value
        End Set
    End Property
    Public Property RangeName() As String
        Get
            Return cRangeName
        End Get
        Set(ByVal Value As String)
            cRangeName = Value
        End Set
    End Property
End Class

Public Class ExcelPack_RptSupervisor
    Dim cColl As New Collection

    Public Sub Add(ByVal SegmentType As Excel.SegmentType, ByVal ClientID As String, ByVal CarrierID As String, ByVal SegmentOffset As Integer, ByVal ReportDate As String, ByVal Location As String, ByVal Sql As String, Optional ByVal SuppressTotalLabel As Boolean = False)
        cColl.Add(New Item(SegmentType, ClientID, CarrierID, SegmentOffset, Location, ReportDate, Sql, SuppressTotalLabel))
    End Sub

    Public ReadOnly Property Coll() As Collection
        Get
            Return cColl
        End Get
    End Property

    Public Class Item
        Private cSegmentType As Excel.SegmentType
        Private cClientID As String
        Private cCarrierID As String
        Private cdt As DataTable
        Private cSegmentOffset As Integer
        Private cReportDate As String
        Private cLocation As String
        Private cSql As String
        Private cSuppressTotalLabel As String

        Public Sub New(ByVal SegmentType As Excel.SegmentType, ByVal ClientID As String, ByVal CarrierID As String, ByVal SegmentOffset As Integer, ByVal ReportDate As String, ByVal Location As String, ByVal Sql As String, Optional ByVal SuppressTotalLabel As Boolean = False)
            cSegmentType = SegmentType
            cClientID = ClientID
            cCarrierID = CarrierID
            cSegmentOffset = SegmentOffset
            cReportDate = ReportDate
            cLocation = Location
            'cdt = dt
            cSql = Sql
            cSuppressTotalLabel = SuppressTotalLabel
        End Sub

        Public ReadOnly Property SegmentType() As Excel.SegmentType
            Get
                Return cSegmentType
            End Get
        End Property

        Public ReadOnly Property ClientID() As String
            Get
                Return cClientID
            End Get
        End Property
        Public ReadOnly Property CarrerID() As String
            Get
                Return cCarrierID
            End Get
        End Property
        Public ReadOnly Property SegmentOffset() As Integer
            Get
                Return cSegmentOffset
            End Get
        End Property
        Public ReadOnly Property ReportDate() As String
            Get
                Return cReportDate
            End Get
        End Property
        Public ReadOnly Property Location() As String
            Get
                Return cLocation
            End Get
        End Property
        Public ReadOnly Property dt() As DataTable
            Get
                Return cdt
            End Get
        End Property
        Public ReadOnly Property Sql() As String
            Get
                Return cSql
            End Get
        End Property
        Public ReadOnly Property SuppressTotalLabel() As Boolean
            Get
                Return cSuppressTotalLabel
            End Get
        End Property
    End Class
End Class

Public Class NotifyFormArgs
    Inherits EventArgs

    Private cSource As SourceEnum
    Private cMessage As String

    Public Enum SourceEnum
        Main = 1
        TablesUpdate = 2
        Rpt_SupvMaster = 3
        Rpt_BVIProduction = 4
        Rpt_EnrCtrMonthly = 5
        Rpt_EnrProductivity = 6
        Excel = 7
        Excel_Generic = 8
        SendEmail = 9
        RptDates = 10
    End Enum

    Public Sub New(ByVal Source As SourceEnum)
        cSource = Source
    End Sub

    Public ReadOnly Property Source() As SourceEnum
        Get
            Return cSource
        End Get
    End Property

    Public Property Message() As String
        Get
            Return cMessage
        End Get
        Set(ByVal Value As String)
            cMessage = Value
        End Set
    End Property
End Class

Public Class CollX
    Inherits System.Collections.CollectionBase
    Private cIsNumber As Boolean
    ' Private Bittem As ListItem

    Public Sub New()
        List.Add(DBNull.Value)
    End Sub

    Public Sub New(ByVal IsNumber As Boolean)
        cIsNumber = IsNumber
    End Sub

    Public Shared Function NewFromList(ByVal ArrayStr As String, ByVal Delimiter As String) As CollX
        Dim i As Integer
        Dim Box As String()
        Dim Coll As New CollX
        Box = ArrayStr.Split("|")
        If Box.GetUpperBound(0) > -1 Then
            For i = 0 To Box.GetUpperBound(0)
                Coll.Assign(Box(i))
            Next
        End If
        Return Coll
    End Function

    Public Overloads ReadOnly Property Count() As Integer
        Get
            Return List.Count - 1
        End Get
    End Property

    Default Public ReadOnly Property Coll(ByVal Idx As Integer) As Object
        Get
            'Return cColl(Idx)
            Return List(Idx).Value
        End Get
    End Property

    Default Public ReadOnly Property Coll(ByVal Key As String) As Object
        Get
            Dim i As Integer
            Dim KeyUpper As String
            KeyUpper = Key.ToUpper
            For i = 1 To List.Count - 1
                If List(i).Key.ToUpper = KeyUpper Then
                    Return List(i).Value
                End If
            Next
            'Dim CollXError As New CollXError("Error #3604: CallX item not found error. Key: " & Key)
            Throw New CollXError("Error #3604: CallX item not found error. Key: " & Key)  'CollXError
        End Get
    End Property

    'Public ReadOnly Property Value(ByVal Key As String) As Object
    '    Get
    '        Dim i As Integer
    '        For i = 1 To List.Count - 1
    '            If List(i).Key = Key Then
    '                'Return List(i).Value
    '                Return List(i).Value
    '            End If
    '        Next
    '        Return Nothing
    '    End Get
    'End Property

    Public ReadOnly Property Key(ByVal Idx As Integer) As String
        Get
            Dim i As Integer
            For i = 1 To List.Count - 1
                If i = Idx Then
                    'Return List(i).Value
                    Return List(i).Key
                End If
            Next
            Return Nothing
        End Get
    End Property

    Public ReadOnly Property DoesKeyExist(ByVal Key As String) As Boolean
        Get
            Dim i As Integer
            Key = Key.ToUpper
            For i = 1 To List.Count - 1
                If List(i).Key.ToUpper = Key Then
                    Return True
                End If
            Next
            Return False
        End Get
    End Property

    Public Sub Assign(ByVal Key As String, ByVal Value As Object)
        Dim i As Integer
        Dim Found As Boolean

        Try

            For i = 1 To List.Count - 1
                If List(i).Key = Key Then
                    List(i).Value = Nothing
                    List(i).Value = Value
                    Found = True
                End If
            Next

            'For Each Item In List
            '    If Item.Key = Key Then
            '        Item.Value = Value
            '        Found = True
            '    End If
            'Next
            If Not Found Then
                List.Add(New KeyValuePair(Key, Value))
            End If

        Catch ex As Exception
            Throw New CollXError("Error #3604: CallX.Assign. Item not found error. Key: " & Key)
        End Try
    End Sub

    Public Sub Assign(ByVal Key_Value As String)
        Dim i As Integer
        Dim Found As Boolean

        Try

            For i = 1 To List.Count - 1
                If List(i).Key = Key_Value Then
                    List(i).Value = Nothing
                    List(i).Value = Key_Value
                    Found = True
                End If
            Next

            'For Each Item In List
            '    If Item.Key = Key Then
            '        Item.Value = Value
            '        Found = True
            '    End If
            'Next
            If Not Found Then
                If cIsNumber Then
                    List.Add(New KeyValuePair(Key_Value, 0))
                Else
                    List.Add(New KeyValuePair(Key_Value, Key_Value))
                End If
            End If

        Catch ex As Exception
            Throw New CollXError("Error #3605: CallX.Assign. " & ex.Message)
        End Try
    End Sub

    Public Overloads Sub RemoveAt(ByVal Index As Integer)
        List.RemoveAt(Index)
    End Sub

    Public Overloads Sub Remove(ByVal Key As String)
        Dim i As Integer
        For i = 1 To List.Count - 1
            If List(i).Key.ToUpper = Key.ToUpper Then
                List.Remove(List(i))
                Exit For
            End If
        Next
    End Sub

    Public Sub ConvertArr(ByRef obj As Object)
        Try

            Dim i As Integer
            For i = 0 To obj.GetUpperBound(0)
                Assign(obj(i))
            Next

        Catch ex As Exception
            Throw New CollXError("Error #3606: CallX.ConvertArr. " & ex.Message)
        End Try
    End Sub

    Public Sub ConvertRow(ByRef dr As DataRow)
        Try

            Dim i As Integer
            For i = 0 To dr.ItemArray.GetUpperBound(0)
                Assign(dr.Table.Columns(i).ColumnName, dr(i))
            Next

        Catch ex As Exception
            Throw New CollXError("Error #3610: CallX.ConvertRow. " & ex.Message)
        End Try
    End Sub

    Public Sub ConvertStr(ByRef Input As String, ByRef Delimiter As String)
        Dim i As Integer
        Dim Box As String()

        Try

            If Input.Length > 0 Then
                If Input.Substring(Input.Length - 1) = Delimiter Then
                    Input = Input.Substring(0, Input.Length - 1)
                End If
                Box = Split(Input, Delimiter)
                For i = 0 To Box.GetUpperBound(0)
                    Assign(Box(i))
                Next
            End If

        Catch ex As Exception
            Throw New CollXError("Error #3607: CallX.ConvertStr. " & ex.Message)
        End Try
    End Sub

    Public Function View() As String()
        Dim i As Integer
        Dim Output(Me.Count) As String
        Dim Val As String

        Try

            For i = 1 To List.Count - 1
                Try
                    Val = List(i).Value
                Catch ex As Exception
                    Val = "<object>"
                End Try
                Output(i) = List(i).Key & "|" & Val
            Next
            Return Output


        Catch ex As Exception
            Throw New CollXError("Error #3608: CallX.View. " & ex.Message)
        End Try
    End Function

    Public Shared Function Clone(ByVal InputColl As CollX) As CollX
        Dim i As Integer
        Dim OutputColl As New CollX

        Try

            For i = 1 To InputColl.Count
                OutputColl.Assign(InputColl.Key(i), InputColl(i))
            Next
            Return OutputColl

        Catch ex As Exception
            Throw New CollXError("Error #3609: CallX.Clone. " & ex.Message)
        End Try
    End Function

    'Public Function Clone(ByRef Item As KeyValuePair) As KeyValuePair
    '    Return New KeyValuePair(Item.Key, Item.Value)
    'End Function

    Public Class KeyValuePair
        Private cKey As String
        Private cValue As Object

        Public Sub New(ByVal Key As String, ByVal Value As Object)
            cKey = Key
            cValue = Value
        End Sub
        Public Property Key() As String
            Get
                Return cKey
            End Get
            Set(ByVal value As String)
                cKey = value
            End Set
        End Property

        Public Property Value() As Object
            Get
                Return cValue
            End Get
            Set(ByVal value As Object)
                cValue = value
            End Set
        End Property
    End Class

    Public Class CollXError
        Inherits Exception
        Private cMessage As String

        Public Sub New(ByVal Message As String)
            cMessage = Message
        End Sub

        Public Overrides ReadOnly Property Message() As String
            Get
                Return cMessage
            End Get
        End Property

        Protected Overrides Sub Finalize()
            MyBase.Finalize()
        End Sub
    End Class
End Class