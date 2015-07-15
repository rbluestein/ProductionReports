'Public Class Rpt_BaseReport
'End Class

Public Class ReportConfig
    Private cCommon As New Common
    Private cReportDate As String
    Private cReportID As Integer
    Private cReportName As String
    Private cTemplateName As String
    Private cDistro As String

    Public Sub New(ByVal ReportID As Integer)
        Dim dt As DataTable

        cReportDate = cCommon.GetServerDateTime.ToString("MM/dd/yyyy")
        'cReportDate = "12/25/2009"

        If ReportID > -1 Then
            dt = cCommon.GetDT("SELECT * FROM Excel_Master WHERE ReportID = " & CType(ReportID, System.String))
            cReportName = dt.Rows(0)("ReportName")
            cDistro = dt.Rows(0)("Distro")
            cTemplateName = cReportName & "_Template.xls"
        End If
    End Sub

    Public ReadOnly Property ReportDate() As String
        Get
            Return cReportDate
        End Get
    End Property

    Public ReadOnly Property ReportName() As String
        Get
            Return cReportName
        End Get
    End Property
    Public ReadOnly Property Distro() As String
        Get
            Return cDistro
        End Get
    End Property
    Public ReadOnly Property TemplateName() As String
        Get
            Return cTemplateName
        End Get
    End Property
    Public ReadOnly Property ReportID() As Integer
        Get
            Return cReportID
        End Get
    End Property
End Class

Public Class EnrollerProductivityReportConfig
    Inherits ReportConfig

    Public Sub New()
        MyBase.New(3)
    End Sub
End Class

Public Class EnrollCenterMonthlyReportConfig
    Inherits ReportConfig

    Public Sub New()
        MyBase.New(2)
    End Sub
End Class

Public Class SupervisorMasterReportConfig
    Inherits ReportConfig

    Private Const cLocationID As String = " u.LocationID IN ('HBG', 'OKC') "
    Private Const cLocationIDStr As String = "HBG/OKC"
    'Results = " u.LocationID IN ('HBG', 'OKC') "
    'Results = " u.LocationID = 'HBG' "
    'Results = " u.LocationID = 'OKC' "

    Public Sub New()
        MyBase.New(1)
    End Sub

    Public ReadOnly Property LocationID() As String
        Get
            Return cLocationID
        End Get
    End Property
    Public ReadOnly Property LocationIDStr() As String
        Get
            Return cLocationIDStr
        End Get
    End Property
End Class

Public Class BVIProductionReportConfig
    Inherits ReportConfig

    Private cReportStartDate As String = "8/12/2009"

    Public Sub New()
        MyBase.New(4)
    End Sub

    Public ReadOnly Property ReportStartDate() As String
        Get
            Return cReportStartDate
        End Get
    End Property
End Class
