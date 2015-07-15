Public Class ProdRpt
    Inherits System.Windows.Forms.Form

    ' Options: 721
    ' Choices: Coalition

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents txtSendEmail As System.Windows.Forms.TextBox
    Friend WithEvents txtRpt_BVIProduction As System.Windows.Forms.TextBox
    Friend WithEvents lblRpt_BVIProduction As System.Windows.Forms.Label
    Friend WithEvents txtRpt_EnrProductivity As System.Windows.Forms.TextBox
    Friend WithEvents lblRpt_EnrProductivity As System.Windows.Forms.Label
    Friend WithEvents txtRpt_EnrCtrMonthly As System.Windows.Forms.TextBox
    Friend WithEvents lblRpt_EnrCtrMonthly As System.Windows.Forms.Label
    Friend WithEvents txtRpt_SupvMaster As System.Windows.Forms.TextBox
    Friend WithEvents lblRpt_SupvMaster As System.Windows.Forms.Label
    Friend WithEvents txtTableUpdate As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtUpdateDates As System.Windows.Forms.TextBox
    Friend WithEvents lblUpdateDates As System.Windows.Forms.LinkLabel
    Friend WithEvents lblTableUpdate As System.Windows.Forms.LinkLabel
    Friend WithEvents lblSendEmail As System.Windows.Forms.LinkLabel
    Friend WithEvents lblGenerateReports As System.Windows.Forms.LinkLabel
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents lblReportDate As System.Windows.Forms.LinkLabel
    Friend WithEvents txtReportDate As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.txtSendEmail = New System.Windows.Forms.TextBox
        Me.txtRpt_BVIProduction = New System.Windows.Forms.TextBox
        Me.lblRpt_BVIProduction = New System.Windows.Forms.Label
        Me.txtRpt_EnrProductivity = New System.Windows.Forms.TextBox
        Me.lblRpt_EnrProductivity = New System.Windows.Forms.Label
        Me.txtRpt_EnrCtrMonthly = New System.Windows.Forms.TextBox
        Me.lblRpt_EnrCtrMonthly = New System.Windows.Forms.Label
        Me.txtRpt_SupvMaster = New System.Windows.Forms.TextBox
        Me.lblRpt_SupvMaster = New System.Windows.Forms.Label
        Me.txtTableUpdate = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtUpdateDates = New System.Windows.Forms.TextBox
        Me.lblUpdateDates = New System.Windows.Forms.LinkLabel
        Me.lblTableUpdate = New System.Windows.Forms.LinkLabel
        Me.lblSendEmail = New System.Windows.Forms.LinkLabel
        Me.lblGenerateReports = New System.Windows.Forms.LinkLabel
        Me.btnClose = New System.Windows.Forms.Button
        Me.lblReportDate = New System.Windows.Forms.LinkLabel
        Me.txtReportDate = New System.Windows.Forms.TextBox
        Me.SuspendLayout()
        '
        'txtSendEmail
        '
        Me.txtSendEmail.BackColor = System.Drawing.Color.White
        Me.txtSendEmail.Location = New System.Drawing.Point(144, 384)
        Me.txtSendEmail.Multiline = True
        Me.txtSendEmail.Name = "txtSendEmail"
        Me.txtSendEmail.ReadOnly = True
        Me.txtSendEmail.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtSendEmail.Size = New System.Drawing.Size(592, 20)
        Me.txtSendEmail.TabIndex = 27
        '
        'txtRpt_BVIProduction
        '
        Me.txtRpt_BVIProduction.BackColor = System.Drawing.Color.White
        Me.txtRpt_BVIProduction.Location = New System.Drawing.Point(144, 248)
        Me.txtRpt_BVIProduction.Multiline = True
        Me.txtRpt_BVIProduction.Name = "txtRpt_BVIProduction"
        Me.txtRpt_BVIProduction.ReadOnly = True
        Me.txtRpt_BVIProduction.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtRpt_BVIProduction.Size = New System.Drawing.Size(592, 32)
        Me.txtRpt_BVIProduction.TabIndex = 25
        '
        'lblRpt_BVIProduction
        '
        Me.lblRpt_BVIProduction.Location = New System.Drawing.Point(32, 248)
        Me.lblRpt_BVIProduction.Name = "lblRpt_BVIProduction"
        Me.lblRpt_BVIProduction.Size = New System.Drawing.Size(104, 23)
        Me.lblRpt_BVIProduction.TabIndex = 24
        Me.lblRpt_BVIProduction.Text = "BVIProduction"
        Me.lblRpt_BVIProduction.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtRpt_EnrProductivity
        '
        Me.txtRpt_EnrProductivity.BackColor = System.Drawing.Color.White
        Me.txtRpt_EnrProductivity.Location = New System.Drawing.Point(144, 336)
        Me.txtRpt_EnrProductivity.Multiline = True
        Me.txtRpt_EnrProductivity.Name = "txtRpt_EnrProductivity"
        Me.txtRpt_EnrProductivity.ReadOnly = True
        Me.txtRpt_EnrProductivity.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtRpt_EnrProductivity.Size = New System.Drawing.Size(592, 32)
        Me.txtRpt_EnrProductivity.TabIndex = 23
        '
        'lblRpt_EnrProductivity
        '
        Me.lblRpt_EnrProductivity.Location = New System.Drawing.Point(32, 336)
        Me.lblRpt_EnrProductivity.Name = "lblRpt_EnrProductivity"
        Me.lblRpt_EnrProductivity.Size = New System.Drawing.Size(104, 23)
        Me.lblRpt_EnrProductivity.TabIndex = 22
        Me.lblRpt_EnrProductivity.Text = "EnrProductivity"
        Me.lblRpt_EnrProductivity.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtRpt_EnrCtrMonthly
        '
        Me.txtRpt_EnrCtrMonthly.BackColor = System.Drawing.Color.White
        Me.txtRpt_EnrCtrMonthly.Location = New System.Drawing.Point(144, 288)
        Me.txtRpt_EnrCtrMonthly.Multiline = True
        Me.txtRpt_EnrCtrMonthly.Name = "txtRpt_EnrCtrMonthly"
        Me.txtRpt_EnrCtrMonthly.ReadOnly = True
        Me.txtRpt_EnrCtrMonthly.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtRpt_EnrCtrMonthly.Size = New System.Drawing.Size(592, 40)
        Me.txtRpt_EnrCtrMonthly.TabIndex = 21
        '
        'lblRpt_EnrCtrMonthly
        '
        Me.lblRpt_EnrCtrMonthly.Location = New System.Drawing.Point(32, 288)
        Me.lblRpt_EnrCtrMonthly.Name = "lblRpt_EnrCtrMonthly"
        Me.lblRpt_EnrCtrMonthly.Size = New System.Drawing.Size(104, 23)
        Me.lblRpt_EnrCtrMonthly.TabIndex = 20
        Me.lblRpt_EnrCtrMonthly.Text = "EnrCtrMonthly"
        Me.lblRpt_EnrCtrMonthly.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtRpt_SupvMaster
        '
        Me.txtRpt_SupvMaster.BackColor = System.Drawing.Color.White
        Me.txtRpt_SupvMaster.Location = New System.Drawing.Point(144, 208)
        Me.txtRpt_SupvMaster.Multiline = True
        Me.txtRpt_SupvMaster.Name = "txtRpt_SupvMaster"
        Me.txtRpt_SupvMaster.ReadOnly = True
        Me.txtRpt_SupvMaster.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtRpt_SupvMaster.Size = New System.Drawing.Size(592, 32)
        Me.txtRpt_SupvMaster.TabIndex = 19
        '
        'lblRpt_SupvMaster
        '
        Me.lblRpt_SupvMaster.Location = New System.Drawing.Point(40, 208)
        Me.lblRpt_SupvMaster.Name = "lblRpt_SupvMaster"
        Me.lblRpt_SupvMaster.Size = New System.Drawing.Size(96, 23)
        Me.lblRpt_SupvMaster.TabIndex = 18
        Me.lblRpt_SupvMaster.Text = "SupvMaster"
        Me.lblRpt_SupvMaster.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtTableUpdate
        '
        Me.txtTableUpdate.BackColor = System.Drawing.Color.White
        Me.txtTableUpdate.Location = New System.Drawing.Point(144, 72)
        Me.txtTableUpdate.Multiline = True
        Me.txtTableUpdate.Name = "txtTableUpdate"
        Me.txtTableUpdate.ReadOnly = True
        Me.txtTableUpdate.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtTableUpdate.Size = New System.Drawing.Size(592, 64)
        Me.txtTableUpdate.TabIndex = 15
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Arial", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(24, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(464, 40)
        Me.Label1.TabIndex = 28
        Me.Label1.Text = "Production Reports Mission Control"
        '
        'txtUpdateDates
        '
        Me.txtUpdateDates.BackColor = System.Drawing.Color.White
        Me.txtUpdateDates.Location = New System.Drawing.Point(144, 48)
        Me.txtUpdateDates.Multiline = True
        Me.txtUpdateDates.Name = "txtUpdateDates"
        Me.txtUpdateDates.ReadOnly = True
        Me.txtUpdateDates.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtUpdateDates.Size = New System.Drawing.Size(592, 20)
        Me.txtUpdateDates.TabIndex = 0
        '
        'lblUpdateDates
        '
        Me.lblUpdateDates.Location = New System.Drawing.Point(40, 48)
        Me.lblUpdateDates.Name = "lblUpdateDates"
        Me.lblUpdateDates.Size = New System.Drawing.Size(96, 23)
        Me.lblUpdateDates.TabIndex = 36
        Me.lblUpdateDates.TabStop = True
        Me.lblUpdateDates.Text = "Get Update Dates"
        Me.lblUpdateDates.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblTableUpdate
        '
        Me.lblTableUpdate.Location = New System.Drawing.Point(40, 72)
        Me.lblTableUpdate.Name = "lblTableUpdate"
        Me.lblTableUpdate.Size = New System.Drawing.Size(96, 23)
        Me.lblTableUpdate.TabIndex = 37
        Me.lblTableUpdate.TabStop = True
        Me.lblTableUpdate.Text = "Update Table"
        Me.lblTableUpdate.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblSendEmail
        '
        Me.lblSendEmail.Location = New System.Drawing.Point(40, 384)
        Me.lblSendEmail.Name = "lblSendEmail"
        Me.lblSendEmail.Size = New System.Drawing.Size(96, 23)
        Me.lblSendEmail.TabIndex = 38
        Me.lblSendEmail.TabStop = True
        Me.lblSendEmail.Text = "Send Email"
        Me.lblSendEmail.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblGenerateReports
        '
        Me.lblGenerateReports.Location = New System.Drawing.Point(40, 184)
        Me.lblGenerateReports.Name = "lblGenerateReports"
        Me.lblGenerateReports.Size = New System.Drawing.Size(96, 16)
        Me.lblGenerateReports.TabIndex = 39
        Me.lblGenerateReports.TabStop = True
        Me.lblGenerateReports.Text = "Generate Reports"
        Me.lblGenerateReports.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(144, 424)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 23)
        Me.btnClose.TabIndex = 41
        Me.btnClose.Text = "Close"
        '
        'lblReportDate
        '
        Me.lblReportDate.Location = New System.Drawing.Point(40, 152)
        Me.lblReportDate.Name = "lblReportDate"
        Me.lblReportDate.Size = New System.Drawing.Size(96, 16)
        Me.lblReportDate.TabIndex = 42
        Me.lblReportDate.TabStop = True
        Me.lblReportDate.Text = "Report Date"
        Me.lblReportDate.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtReportDate
        '
        Me.txtReportDate.BackColor = System.Drawing.Color.White
        Me.txtReportDate.Location = New System.Drawing.Point(144, 152)
        Me.txtReportDate.Name = "txtReportDate"
        Me.txtReportDate.Size = New System.Drawing.Size(576, 20)
        Me.txtReportDate.TabIndex = 43
        '
        'ProdRpt
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(784, 462)
        Me.Controls.Add(Me.txtReportDate)
        Me.Controls.Add(Me.txtUpdateDates)
        Me.Controls.Add(Me.txtSendEmail)
        Me.Controls.Add(Me.txtRpt_BVIProduction)
        Me.Controls.Add(Me.txtRpt_EnrProductivity)
        Me.Controls.Add(Me.txtRpt_EnrCtrMonthly)
        Me.Controls.Add(Me.txtRpt_SupvMaster)
        Me.Controls.Add(Me.txtTableUpdate)
        Me.Controls.Add(Me.lblReportDate)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.lblGenerateReports)
        Me.Controls.Add(Me.lblSendEmail)
        Me.Controls.Add(Me.lblTableUpdate)
        Me.Controls.Add(Me.lblUpdateDates)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.lblRpt_BVIProduction)
        Me.Controls.Add(Me.lblRpt_EnrProductivity)
        Me.Controls.Add(Me.lblRpt_EnrCtrMonthly)
        Me.Controls.Add(Me.lblRpt_SupvMaster)
        Me.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "ProdRpt"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

#Region " Declarations "
    Private cEnviro As Enviro
    Private cCommon As Common
    Private cReport As Report
    Private Const Sleeptime As Integer = 4000
    Private cAttachmentColl As New Collection
#End Region

#Region " Display Messages "
    Private Sub DisplayMessage(ByRef e As NotifyFormArgs)
        Dim ctl As TextBox

        Select Case e.Source

            Case NotifyFormArgs.SourceEnum.TablesUpdate
                ctl = txtTableUpdate
            Case NotifyFormArgs.SourceEnum.Rpt_SupvMaster
                ctl = txtRpt_SupvMaster
            Case NotifyFormArgs.SourceEnum.Rpt_BVIProduction
                ctl = txtRpt_BVIProduction
            Case NotifyFormArgs.SourceEnum.Rpt_EnrCtrMonthly
                ctl = txtRpt_EnrCtrMonthly
            Case NotifyFormArgs.SourceEnum.Rpt_EnrProductivity
                ctl = txtRpt_EnrProductivity
            Case NotifyFormArgs.SourceEnum.SendEmail
                ctl = txtSendEmail
            Case NotifyFormArgs.SourceEnum.RptDates
                ctl = txtUpdateDates
        End Select
        DisplayMessage2(ctl, e.Message)
    End Sub

    Private Sub DisplayMessage2(ByRef ctl As TextBox, ByVal Message As String)
        ctl.Text = Message
        ctl.Refresh()
        Me.Refresh()
        Application.DoEvents()
        System.Diagnostics.Debug.WriteLine(Message)
    End Sub


#End Region

#Region " Page load "
    Private Sub ProdRpt_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        cEnviro = New Enviro
        gEnviro = cEnviro
        cCommon = New Common
        cReport = New Report

        ' ___ Initialize Enviro values
        cEnviro.DBHost = "192.168.1.10"
        cEnviro.LogFileFullPath = cEnviro.GetAppPath & "\ProductionReportsLog.txt"

        txtUpdateDates.Text = "Not started."
        txtUpdateDates.SelectionStart = 0
        txtUpdateDates.SelectionLength = 0
        txtTableUpdate.Text = "Not started"
        txtRpt_SupvMaster.Text = "Not started"
        txtRpt_EnrCtrMonthly.Text = "Not started"
        txtRpt_EnrProductivity.Text = "Not started"
        txtRpt_BVIProduction.Text = "Not started"
        txtSendEmail.Text = "Not started"

        If Not cCommon.TestServerConnection(5) Then
            lblUpdateDates.Enabled = False
            lblTableUpdate.Enabled = False
            lblReportDate.Enabled = False
            lblGenerateReports.Enabled = False
            lblSendEmail.Enabled = False
            MessageBox.Show("Production Reports is unable to connect to server.")
        End If
        Me.Text = "Production reports v" & cEnviro.VersionNumber
    End Sub
#End Region

#Region " Unused "
    'Private Sub NavCtl(ByVal Source As String, ByVal Info As String)


    '    'lblUpdateDates.Enabled = True
    '    'lblTableUpdate.Enabled = True
    '    'lblReportDate.Enabled = True
    '    'lblGenerateReports.Enabled = True
    '    'lblSendEmail.Enabled = True

    '    'lblTableUpdate.Enabled = False
    '    'lblReportDate.Enabled = False
    '    'lblGenerateReports.Enabled = False
    '    'lblSendEmail.Enabled = False


    '    Select Case Source
    '        Case "FormLoad"
    '            Select Case Info
    '                Case "yes"
    '                    lblUpdateDates.Enabled = True
    '                Case "no"
    '                    lblUpdateDates.Enabled = False
    '            End Select
    '            lblTableUpdate.Enabled = False
    '            lblReportDate.Enabled = False
    '            lblGenerateReports.Enabled = False
    '            lblSendEmail.Enabled = True

    '        Case "GetUpdateDates"

    '        Case "UpdateTable"

    '        Case "GenerateReports"

    '        Case "SendEmail"

    '    End Select
    'End Sub
#End Region

#Region " Button clicked "
#Region " Get update dates"
    Private Sub lblUpdateDates_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lblUpdateDates.LinkClicked

        'Dim datafix As New Datafix
        'datafix.AddBVIFLegal()
        'Stop

        Dim TablesUpdate As New TablesUpdate(TablesUpdate.CallTypeEnum.UpdateDates)
        Dim ee As New NotifyFormArgs(NotifyFormArgs.SourceEnum.RptDates)
        Dim Results As Results

        ee.Message = "Looking up dates"
        DisplayMessage(ee)
        TablesUpdate = New TablesUpdate(TablesUpdate.CallTypeEnum.UpdateDates)
        AddHandler TablesUpdate.NotifyForm, AddressOf DisplayMessage
        Results = TablesUpdate.Init
        ee.Message = Results.Message
        DisplayMessage(ee)
    End Sub
#End Region

#Region " Perform tables update clicked "
    Private Sub lblTableUpdate_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lblTableUpdate.LinkClicked
        Dim Results As Results

        If txtUpdateDates.Text = "Not started." Then
            txtTableUpdate.Text = "Update dates not provided."
            Exit Sub
        End If

        Dim TablesUpdate As New TablesUpdate(TablesUpdate.CallTypeEnum.FullUpdate)
        AddHandler TablesUpdate.NotifyForm, AddressOf DisplayMessage
        Results = TablesUpdate.Init
        If Results.Success Then
            DisplayMessage2(txtTableUpdate, "Done.")
        Else
            DisplayMessage2(txtTableUpdate, Results.Message)
        End If
    End Sub
#End Region

#Region " Get report date clicked "
    Private Sub lblReportDate_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lblReportDate.LinkClicked
        Dim ReportConfig As New ReportConfig(-1)
        DisplayMessage2(txtReportDate, cCommon.GetReportDate(ReportConfig.ReportDate, Nothing))
    End Sub
#End Region

#Region " Generate reports clicked "
    Private Sub lblGenerateReports_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lblGenerateReports.LinkClicked
        Dim MyResults As New Results
        Dim Results As Results
        Dim Rpt_SupvMaster As Rpt_SupvMaster
        Dim Rpt_EnrCtrMonthly As Rpt_EnrCtrMonthly
        Dim Rpt_EnrProductivity As Rpt_EnrProductivity
        Dim Rpt_BVIProduction As Rpt_BVIProduction

        Try

            If Not IsDate(txtReportDate.Text) Then
                txtReportDate.Text = "Report date not provided."
                Exit Sub
            End If

            ' ___ Attachments collection
            cAttachmentColl = New Collection

            ' ___ Set success default value
            MyResults.Success = True

            ' ___ Set report run datetime
            cEnviro.ReportDateTime = cCommon.GetServerDateTime.ToString("yyyyMMdd_HHmmss")

            ' ___ Supervisor report
            'If 0 = 0 Then
            '    DisplayMessage2(txtRpt_SupvMaster, "Building report...")
            '    Rpt_SupvMaster = New Rpt_SupvMaster
            '    AddHandler Rpt_SupvMaster.NotifyForm, AddressOf DisplayMessage
            '    Results = Rpt_SupvMaster.Init(txtReportDate.Text)
            '    If Results.Success Then
            '        cAttachmentColl.Add(Rpt_SupvMaster.OutputFullPath, "Rpt_SupvMaster")
            '        DisplayMessage2(txtRpt_SupvMaster, "Done.")
            '    Else
            '        MyResults.Success = False
            '        DisplayMessage2(txtRpt_SupvMaster, Results.Message)
            '    End If
            'End If


            ' ___ BVI Production  Report
            If 0 = 0 Then
                If MyResults.Success Then
                    DisplayMessage2(txtRpt_BVIProduction, "Building report...")
                    Rpt_BVIProduction = New Rpt_BVIProduction
                    AddHandler Rpt_BVIProduction.NotifyForm, AddressOf DisplayMessage
                    Results = Rpt_BVIProduction.Init(txtReportDate.Text)
                    If Results.Success Then
                        cAttachmentColl.Add(Rpt_BVIProduction.OutputFullPath, "Rpt_BVIProduction")
                        DisplayMessage2(txtRpt_BVIProduction, "Done.")
                    Else
                        MyResults.Success = False
                        DisplayMessage2(txtRpt_BVIProduction, Results.Message)
                    End If
                End If
            End If


            ' ___ Enrollment Center Monthly  Report
            If 0 = 1 Then
                If MyResults.Success Then
                    DisplayMessage2(txtRpt_EnrCtrMonthly, "Building report...")
                    Rpt_EnrCtrMonthly = New Rpt_EnrCtrMonthly
                    Results = Rpt_EnrCtrMonthly.Init
                    If Results.Success Then
                        cAttachmentColl.Add(Rpt_EnrCtrMonthly.OutputFullPath)
                        DisplayMessage2(txtRpt_EnrCtrMonthly, "Done.")

                    Else
                        MyResults.Success = False
                        DisplayMessage2(txtRpt_EnrCtrMonthly, Results.Message)
                    End If
                End If
            End If

            ' ___ Enroller Productivity  Report
            If 0 = 1 Then
                If MyResults.Success Then
                    DisplayMessage2(txtRpt_EnrProductivity, "Building report...")
                    Rpt_EnrProductivity = New Rpt_EnrProductivity
                    Results = Rpt_EnrProductivity.Init
                    If Results.Success Then
                        cAttachmentColl.Add(Rpt_EnrProductivity.OutputFullPath)
                        DisplayMessage2(txtRpt_EnrProductivity, "Done.")
                    Else
                        MyResults.Success = False
                        DisplayMessage2(txtRpt_EnrProductivity, Results.Message)
                    End If
                End If
            End If

        Catch ex As Exception
            cReport.Report("Main #100 " & ex.Message, Report.ReportTypeEnum.LogError)
        End Try
    End Sub
#End Region

#Region " Send Email clicked "
    Private Sub lblSendEmail_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lblSendEmail.LinkClicked
        Dim i As Integer
        Dim DirInfo As System.IO.DirectoryInfo
        Dim FileInfo As System.IO.FileInfo()
        Dim DirPath As String
        Dim FullPath As String
        Dim Box() As String
        Dim StreamReader As System.IO.StreamReader
        Dim FileText As String = String.Empty
        Dim LastReportDate As String
        Dim ee As New NotifyFormArgs(NotifyFormArgs.SourceEnum.SendEmail)
        Dim Results As Results
        Dim Coll As New Collection
        Dim CurFile As String
        Dim CurDate As DateTime
        Dim DialogResult As DialogResult

        Try

            If Not IsDate(txtReportDate.Text) Then
                txtSendEmail.Text = "Report date not provided."
                Exit Sub
            End If

            ee.Message = "Sending email"
            DisplayMessage(ee)

            ' ___ AddlData
            DirPath = cEnviro.GetAppPath & "\TempData\AddlData"
            DirInfo = New System.IO.DirectoryInfo(DirPath)
            FileInfo = DirInfo.GetFiles

            If FileInfo.GetUpperBound(0) > -1 Then
                If (MessageBox.Show("Production Reports will email files found in AddlData. Do you wish to proceed?", Nothing, MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2)) = DialogResult.Cancel Then
                    Exit Sub
                End If
            End If

            For i = 0 To FileInfo.GetUpperBound(0)
                FullPath = DirPath & "\" & FileInfo(i).Name
                Box = Split(FullPath, ".")
                Select Case Box(Box.GetUpperBound(0))
                    Case "xls"
                        cAttachmentColl.Add(FullPath)
                    Case "txt"
                        StreamReader = New System.IO.StreamReader(FullPath)
                        FileText = StreamReader.ReadToEnd
                End Select
            Next


            ' ___ Ron Kleiman
            LastReportDate = "Production report -- " & CType(txtReportDate.Text, System.DateTime).ToString("M/d/yyyy")
            'Results = cReport.SendEmail("rkleiman@benefitvision.com", "automail@benefitvision.com", "jresor@benefitvision.com;iyacht@benefitvision.com; jkleiman@benefitvision.com; cschwartz@benefitvision.com; rbluestein@benefitvision.com", LastReportDate, FileText, cAttachmentColl)
            Results = cReport.SendEmail("rkleiman@benefitvision.com", "jkleiman@benefitvision.com", "jresor@benefitvision.com;iyacht@benefitvision.com; cschwartz@benefitvision.com; rbluestein@benefitvision.com", LastReportDate, FileText, cAttachmentColl)
            If Results.Success Then
                ee.Message = "Email sent to Ron Kleiman. "
            Else
                ee.Message = "Ron Kleiman: " & Results.Message & ". "
            End If
            DisplayMessage(ee)


            ' ___ Rick Smith
            Try
                Coll.Add(cAttachmentColl("Rpt_SupvMaster"))
            Catch ex As Exception
                For i = 1 To cAttachmentColl.Count
                    If InStr(cAttachmentColl(i), "ProdRpts_Rpt_SupvMaster") > 0 Then
                        Coll.Add(cAttachmentColl(i))
                    End If
                Next
            End Try


            FileText = String.Empty
            LastReportDate = "Supervisor master report -- " & CType(txtReportDate.Text, System.DateTime).ToString("M/d/yyyy")
            Results = cReport.SendEmail("jkleiman@benefitvision.com", "automail@benefitvision.com", Nothing, LastReportDate, FileText, Coll)
            If Results.Success Then
                ee.Message &= "Email sent to Jen Kleiman. "
            Else
                ee.Message &= "Jen Kleiman: " & Results.Message & ". "
            End If
            DisplayMessage(ee)



            ' ___ Art MacAuley
            DirPath = cEnviro.GetAppPath & "\ArtMacCauley"
            DirInfo = New System.IO.DirectoryInfo(DirPath)
            FileInfo = DirInfo.GetFiles
            If FileInfo.GetUpperBound(0) > 0 Then
                CurFile = FileInfo(0).Name
                CurDate = FileInfo(0).CreationTime
                For i = 1 To FileInfo.GetUpperBound(0)
                    If FileInfo(i).CreationTime > CurDate Then
                        CurFile = FileInfo(i).Name
                        CurDate = FileInfo(i).CreationTime
                    End If
                Next

                DialogResult = MessageBox.Show("Send Choices production report dated " & CurDate.ToString("M/d/yyyy") & " to Art MacCauley?", "Art MacCauley", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)

                If DialogResult.Yes Then
                    LastReportDate = "Choices production report -- " & Date.Now.ToString("M/d/yyyy")
                    Coll = New Collection
                    Coll.Add(DirPath & "\" & CurFile)
                    Results = cReport.SendEmail("arthurp227@bellsouth.net", "automail@benefitvision.com", Nothing, LastReportDate, FileText, Coll)
                    If Results.Success Then
                        ee.Message &= "Email sent to Art MacCauley."
                    Else
                        ee.Message &= "ArtMacCauley: " & Results.Message & "."
                    End If
                    DisplayMessage(ee)
                End If

            End If






            For i = 1 To cAttachmentColl.Count
                If InStr(cAttachmentColl(i), "Rpt_SupvMaster") = 0 Then

                End If
            Next

            'ProdRpts_Rpt_SupvMaster_20100825_190706.xls

        Catch ex As Exception
            Throw New Exception("Error #1810: Main PrepareEmail. " & ex.Message)
        Finally
            Try
                StreamReader.Close()
            Catch
            End Try
        End Try
    End Sub
#End Region

#Region " Close app clicked "
    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        cCommon.ExitApplication()
    End Sub
#End Region
#End Region
End Class