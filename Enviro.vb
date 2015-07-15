Public Class Enviro
#Region " Declarations "
    Private cDBHost As String
    Private cLogFileFullPath As String
    Private cReportTablesUpdateLogFileFullPath As String
    Private cReportTablesUpdateColl As New Collection
    Private cReportDateTime As String
    Private cVersionNumber As String = "1.74"
#End Region

#Region " Constants "
    Private Const cConnStringTemplate As String = "user id=BVI_SQL_SERVER;password=noisivtifeneb;database=|;server="
    'Private Const cConnStringTemplate As String = "user id=BVI_SQL_SERVER;password=noisivtifeneb;timeout=20;database=|;server="
    Private Const cDBName As String = "ProjectReports"
    Private Const cIncludeTodayInUpdate As Boolean = False
    Private Const cAllowTableUpdatePreviousYear As Boolean = False
    Private Const cAllowUpdateRptBVIProductionPreviousYear As Boolean = False
#End Region

#Region " Properties "
    Public ReadOnly Property DBName() As String
        Get
            Return cDBName
        End Get
    End Property

    Public Property DBHost() As String
        Get
            Return cDBHost
        End Get
        Set(ByVal Value As String)
            cDBHost = Value
        End Set
    End Property

    Public Property ReportDateTime() As String
        Get
            Return cReportDateTime
        End Get
        Set(ByVal Value As String)
            cReportDateTime = Value
        End Set
    End Property

    Public Property LogFileFullPath() As String
        Get
            Return cLogFileFullPath
        End Get
        Set(ByVal Value As String)
            cLogFileFullPath = Value
        End Set
    End Property

    Public ReadOnly Property IncludeTodayInUpdate() As Boolean
        Get
            Return cIncludeTodayInUpdate
        End Get
    End Property

    Public ReadOnly Property AllowTableUpdatePreviousYear() As Boolean
        Get
            Return cAllowTableUpdatePreviousYear
        End Get
    End Property
    Public ReadOnly Property AllowUpdateRptBVIProductionPreviousYear() As Boolean
        Get
            Return cAllowUpdateRptBVIProductionPreviousYear
        End Get
    End Property
    Public ReadOnly Property VersionNumber() As String
        Get
            Return cVersionNumber
        End Get
    End Property
#End Region

#Region " Methods "
    Public Function ConnectionString() As String
        Return Replace(cConnStringTemplate, "|", cDBName) & cDBHost
    End Function

    Public Function ConnectionString(ByVal DBName As String) As String
        Return Replace(cConnStringTemplate, "|", DBName) & cDBHost
    End Function

    Public Function ConnectionString(ByVal DBHost As String, ByVal DBName As String) As String
        Return Replace(cConnStringTemplate, "|", DBName) & DBHost
    End Function

    Public Function TestIDSelect(ByVal DBName As String, ByVal EmpID As String) As String
        Return " AND dbo.ufn_IsTestID(" & DBName & ", " & EmpID & ") = 0 "
    End Function

    Public Function GetAppPath() As String
        Dim i As Integer
        Dim FullPath As String
        Dim FullPath_() As String
        Dim Results As String
        Dim Count As Integer

        FullPath = System.IO.Directory.GetCurrentDirectory()
        FullPath_ = Split(FullPath, "\")

        Count = FullPath_.GetUpperBound(0)

        If FullPath_(FullPath_.GetUpperBound(0)) = "bin" Then
            Count -= 1
        End If

        For i = 0 To Count
            Results = Results & FullPath_(i) & "\"
        Next
        Results = Results.Substring(0, Results.Length - 1)
        Return Results
    End Function
#End Region
End Class