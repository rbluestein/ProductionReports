'Public Class Form1
'    Inherits System.Windows.Forms.Form

'    Private cFullPath As String
'    Private cSelectionStart As Integer
'    Private cSelectionLength As Integer
'    Private cExcelExecutingInd As Boolean

'    Private cEnviro As New Enviro

'#Region " Windows Form Designer generated code "

'    Public Sub New()
'        MyBase.New()

'        'This call is required by the Windows Form Designer.
'        InitializeComponent()

'        'Add any initialization after the InitializeComponent() call

'    End Sub

'    'Form overrides dispose to clean up the component list.
'    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
'        If disposing Then
'            If Not (components Is Nothing) Then
'                components.Dispose()
'            End If
'        End If
'        MyBase.Dispose(disposing)
'    End Sub

'    'Required by the Windows Form Designer
'    Private components As System.ComponentModel.IContainer

'    'NOTE: The following procedure is required by the Windows Form Designer
'    'It can be modified using the Windows Form Designer.  
'    'Do not modify it using the code editor.
'    Friend WithEvents btnGo As System.Windows.Forms.Button
'    Friend WithEvents txtSql As System.Windows.Forms.TextBox
'    Friend WithEvents btnNewProd As System.Windows.Forms.Button
'    Friend WithEvents btnNewTest As System.Windows.Forms.Button
'    Friend WithEvents btnFormat As System.Windows.Forms.Button
'    Friend WithEvents ddDatabase As System.Windows.Forms.ComboBox
'    Friend WithEvents ddServer As System.Windows.Forms.ComboBox
'    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
'        Me.btnGo = New System.Windows.Forms.Button
'        Me.txtSql = New System.Windows.Forms.TextBox
'        Me.btnNewProd = New System.Windows.Forms.Button
'        Me.btnNewTest = New System.Windows.Forms.Button
'        Me.btnFormat = New System.Windows.Forms.Button
'        Me.ddDatabase = New System.Windows.Forms.ComboBox
'        Me.ddServer = New System.Windows.Forms.ComboBox
'        Me.SuspendLayout()
'        '
'        'btnGo
'        '
'        Me.btnGo.Location = New System.Drawing.Point(8, 152)
'        Me.btnGo.Name = "btnGo"
'        Me.btnGo.Size = New System.Drawing.Size(72, 23)
'        Me.btnGo.TabIndex = 2
'        Me.btnGo.Text = "Go"
'        '
'        'txtSql
'        '
'        Me.txtSql.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
'        Me.txtSql.Location = New System.Drawing.Point(8, 32)
'        Me.txtSql.Multiline = True
'        Me.txtSql.Name = "txtSql"
'        Me.txtSql.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
'        Me.txtSql.Size = New System.Drawing.Size(336, 112)
'        Me.txtSql.TabIndex = 1
'        Me.txtSql.Text = ""
'        '
'        'btnNewProd
'        '
'        Me.btnNewProd.Location = New System.Drawing.Point(272, 152)
'        Me.btnNewProd.Name = "btnNewProd"
'        Me.btnNewProd.Size = New System.Drawing.Size(72, 23)
'        Me.btnNewProd.TabIndex = 5
'        Me.btnNewProd.Text = "New"
'        '
'        'btnNewTest
'        '
'        Me.btnNewTest.Location = New System.Drawing.Point(192, 152)
'        Me.btnNewTest.Name = "btnNewTest"
'        Me.btnNewTest.Size = New System.Drawing.Size(72, 23)
'        Me.btnNewTest.TabIndex = 4
'        Me.btnNewTest.Text = "New Test"
'        Me.btnNewTest.Visible = False
'        '
'        'btnFormat
'        '
'        Me.btnFormat.Location = New System.Drawing.Point(88, 152)
'        Me.btnFormat.Name = "btnFormat"
'        Me.btnFormat.Size = New System.Drawing.Size(72, 23)
'        Me.btnFormat.TabIndex = 3
'        Me.btnFormat.Text = "Format"
'        '
'        'ddDatabase
'        '
'        Me.ddDatabase.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
'        Me.ddDatabase.Location = New System.Drawing.Point(8, 8)
'        Me.ddDatabase.Name = "ddDatabase"
'        Me.ddDatabase.Size = New System.Drawing.Size(121, 21)
'        Me.ddDatabase.TabIndex = 6
'        '
'        'ddServer
'        '
'        Me.ddServer.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
'        Me.ddServer.Location = New System.Drawing.Point(136, 8)
'        Me.ddServer.Name = "ddServer"
'        Me.ddServer.Size = New System.Drawing.Size(80, 21)
'        Me.ddServer.TabIndex = 7
'        '
'        'Form1
'        '
'        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
'        Me.ClientSize = New System.Drawing.Size(368, 190)
'        Me.Controls.Add(Me.ddServer)
'        Me.Controls.Add(Me.ddDatabase)
'        Me.Controls.Add(Me.btnFormat)
'        Me.Controls.Add(Me.btnNewTest)
'        Me.Controls.Add(Me.btnNewProd)
'        Me.Controls.Add(Me.txtSql)
'        Me.Controls.Add(Me.btnGo)
'        Me.MinimumSize = New System.Drawing.Size(360, 216)
'        Me.Name = "Form1"
'        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Show
'        Me.Text = "New Excel Big 3.0"
'        Me.ResumeLayout(False)

'    End Sub

'#End Region

'    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
'        Dim PassedInArgs As String()
'        Dim dt As DataTable
'        Dim i As Integer
'        Dim ReadStream As System.IO.StreamReader
'        Dim Box() As String
'        Dim IniData As String
'        Dim DBName As String
'        Dim ServerName As String

'        cFullPath = System.IO.Directory.GetCurrentDirectory() & "\NewExcelBig.ini"

'        'PassedInArgs = System.Environment.GetCommandLineArgs
'        'MessageBox.Show(PassedInArgs.GetUpperBound(0))

'        Try

'            'If PassedInArgs.GetUpperBound(0) = 0 Then
'            '    Master.IPAddress = "192.168.1.15"
'            'Else

'            '    If PassedInArgs(1) = "prod" Then
'            '        Master.IPAddress = "192.168.1.10"
'            '    Else
'            '        Master.IPAddress = "192.168.1.15"
'            '    End If
'            'End If

'            'If Master.IPAddress = "192.168.1.10" Then
'            '    Me.Text = "New Excel hbg-prod"
'            'ElseIf Master.IPAddress = "192.168.1.15" Then
'            '    Me.Text = "New Excel hbg-tst"
'            'Else
'            '    MessageBox.Show("Unable to determine environment")
'            'End If

'            ' ___ Read the ini file
'            If System.IO.File.Exists(cFullPath) Then
'                ReadStream = New System.IO.StreamReader(cFullPath)
'                IniData &= ReadStream.ReadToEnd
'                ReadStream.Close()
'                Box = Split(IniData, "|")
'                ServerName = Box(0)
'                DBName = Box(1)
'            End If

'            ' ___ Populate the server dropdown. Select item if available. Set the Master.IPAddress in order to obtain the database names.
'            ddServer.Items.Add("hbg-tst")
'            ddServer.Items.Add("hbg-sql")
'            ddServer.Items.Add("okc-sql")
'            If IniData = Nothing Then
'                cenviro.serveraddress = "192.168.1.10"
'            Else
'                ddServer.Text = ServerName
'                Select Case ServerName
'                    Case "hbg-tst"
'                        cenviro.serveraddress = "192.168.1.15"
'                    Case "hbg-sql"
'                        cenviro.serveraddress = "192.168.1.10"
'                    Case "okc-sql"
'                        cenviro.serveraddress = "192.168.2.10"
'                End Select
'            End If


'            ' ___ Populate the database dropdown. Select item if available.
'            dt = cCommon.GetDT("SELECT * FROM master..sysdatabases ORDER BY NAME")
'            For i = 0 To dt.Rows.Count - 1
'                ddDatabase.Items.Add(dt.Rows(i)(0))
'            Next
'            If IniData <> Nothing Then
'                ddDatabase.Text = DBName
'            End If



'        Catch ex As Exception
'            cReport.Report(Report.ReportTypeEnum.Error, "Error #330: Form_Load " & ex.Message)
'        End Try
'    End Sub


'    Private Sub btnGo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGo.Click
'        Try
'            If txtSql.Text.Length > 0 Then
'                txtSql.Text = Trim(txtSql.Text)
'                Go()
'            End If
'        Catch ex As Exception
'            cReport.Report(Report.ReportTypeEnum.Error, "Error #331: btnGo_Click " & ex.Message)
'        End Try
'    End Sub


'    Private Sub Go()
'        Dim Results As Results
'        Dim QueryPack As DBase.QueryPack
'        Dim Sql As String

'        Try

'            Me.Cursor = Cursors.WaitCursor
'            Dim Excel As New Excel
'            Dim dt As DataTable

'            Select Case ddServer.SelectedItem
'                Case "hbg-tst"
'                    Master.IPAddress = "192.168.1.15"
'                Case "hbg-sql"
'                    Master.IPAddress = "192.168.1.10"
'                Case "okc-sql"
'                    Master.IPAddress = "192.168.2.10"
'            End Select

'            txtSql.Text = Trim(txtSql.Text)

'            cSelectionLength = txtSql.SelectionLength
'            cSelectionStart = txtSql.SelectionStart

'            If cSelectionLength > 0 Then
'                Sql = txtSql.Text.Substring(cSelectionStart, cSelectionLength)
'            Else
'                Sql = txtSql.Text
'            End If
'            Sql = "USE [" & ddDatabase.SelectedItem & "] " & Sql

'            cExcelExecutingInd = True

'            QueryPack = cCommon.GetDTWithQueryPack(Sql)
'            If Not QueryPack.Success Then
'                MessageBox.Show("Database error:" & vbCrLf & QueryPack.TechErrMsg)
'            End If

'            If QueryPack.Success Then
'                Results = Excel.ExportToExcel(QueryPack.dt, Nothing)
'                If Not Results.Success Then
'                    Select Case txtSql.Text.ToLower.Substring(0, 6)
'                        Case "select"
'                            MessageBox.Show("Excel error:" & vbCrLf & Results.Msg)
'                        Case "insert"
'                            If Results.Msg = "Exception from HRESULT: 0x800A03EC." Then
'                                MessageBox.Show("INSERT query completed. Check for accuracy.")
'                            End If
'                        Case "update"
'                            If Results.Msg = "Exception from HRESULT: 0x800A03EC." Then
'                                MessageBox.Show("UPDATE query completed. Check for accuracy.")
'                            End If
'                    End Select
'                End If
'            End If


'            Dim WriteStream As New System.IO.StreamWriter(cFullPath)
'            WriteStream.Write(ddServer.SelectedItem & "|" & ddDatabase.SelectedItem)
'            WriteStream.Close()



'            'txtSql.SelectionLength = SelectionLength
'            'txtSql.SelectionStart = SelectionStart

'            Me.Cursor = Cursors.Default

'            'Me.Refresh()



'        Catch ex As Exception
'            cReport.Report(Report.ReportTypeEnum.Error, "Error #332: Go " & ex.Message)
'        End Try
'    End Sub

'    Private Sub btnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNewProd.Click
'        Try
'            'Shell("NewExcelBig.exe prod", AppWinStyle.NormalFocus)
'            Shell("NewExcelBig.exe", AppWinStyle.NormalFocus)
'        Catch ex As Exception
'            cReport.Report(Report.ReportTypeEnum.Error, "Error #333: btnNew_Click " & ex.Message)
'        End Try
'    End Sub

'    Private Sub btnNewTest_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNewTest.Click
'        Try
'            Shell("NewExcelBig.exe test", AppWinStyle.NormalFocus)
'        Catch ex As Exception
'            cReport.Report(Report.ReportTypeEnum.Error, "Error #334: btnNewTest_Click " & ex.Message)
'        End Try
'    End Sub

'    Private Sub FrmMain_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Resize
'        txtSql.Height = Me.Height - 104
'        txtSql.Width = Me.Width - 24

'        btnGo.Top = Me.Height - 65
'        btnFormat.Top = Me.Height - 65
'        btnNewTest.Top = Me.Height - 65
'        btnNewProd.Top = Me.Height - 65

'        'SSTab1.Height = Me.Height - 152
'        'treProductionProgs.Height = SSTab1.Height - 32
'        'treTestProgs.Height = SSTab1.Height - 32
'        'Panel1.Top = SSTab1.Top + SSTab1.Height
'    End Sub

'    Private Sub btnFormat_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFormat.Click
'        'txtSql.Text = Replace(txtSql.Text, vbCrLf, " ", 1, -1, CompareMethod.Text)

'        txtSql.Text = Replace(txtSql.Text, vbCrLf, "")
'        txtSql.Text = Replace(txtSql.Text, """", "" & vbCrLf, 1, -1, CompareMethod.Text)
'        ' txtSql.Text = Replace(txtSql.Text, "SELECT", vbCrLf & "SELECT" & vbCrLf, 1, -1, CompareMethod.Text)
'        txtSql.Text = Replace(txtSql.Text, " SELECT", vbCrLf & "SELECT", 1, -1, CompareMethod.Text)
'        txtSql.Text = Replace(txtSql.Text, "FROM", vbCrLf & "FROM", 1, -1, CompareMethod.Text)
'        txtSql.Text = Replace(txtSql.Text, "INNER", vbCrLf & "INNER", 1, -1, CompareMethod.Text)
'        txtSql.Text = Replace(txtSql.Text, "WHERE", vbCrLf & "WHERE", 1, -1, CompareMethod.Text)
'        txtSql.Text = Replace(txtSql.Text, "ORDER BY", vbCrLf & "ORDER BY", 1, -1, CompareMethod.Text)
'        txtSql.Text = Replace(txtSql.Text, "AND ", "AND" & vbCrLf, 1, -1, CompareMethod.Text)
'    End Sub




'    'Private Sub Form1_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated
'    '        txtSql.SelectionLength = cSelectionLength
'    '        txtSql.SelectionStart = cSelectionStart
'    'End Sub

'    'Private Sub Form1_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.MouseDown
'    '    txtSql.SelectionLength = cSelectionLength
'    '    txtSql.SelectionStart = cSelectionStart
'    'End Sub

'    'Private Sub txtSql_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSql.GotFocus
'    '    txtSql.SelectionLength = cSelectionLength
'    '    txtSql.SelectionStart = cSelectionStart
'    'End Sub

'    'Private Sub txtSql_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles txtSql.MouseUp
'    '    txtSql.SelectionLength = cSelectionLength
'    '    txtSql.SelectionStart = cSelectionStart
'    'End Sub

'    'Private Sub Form1_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseDown
'    '    txtSql.SelectionLength = cSelectionLength
'    '    txtSql.SelectionStart = cSelectionStart
'    'End Sub

'    Private Sub ExcelReturn()
'        If cExcelExecutingInd Then

'            'txtSql.Text = "abcdefghijklmnopqrstuvwxyz"
'            'txtSql.SelectionLength = 4
'            'txtSql.SelectionStart = 3

'            cExcelExecutingInd = False
'            txtSql.SelectionLength = cSelectionLength
'            txtSql.SelectionStart = cSelectionStart
'            Me.Focus()
'            Me.Refresh()
'        End If
'    End Sub

'    Private Sub Form1_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated
'        ExcelReturn()
'    End Sub

'    Private Sub Form1_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Enter
'        ExcelReturn()
'    End Sub


'    Protected Overrides Sub OnGotFocus(ByVal e As System.EventArgs)
'        ExcelReturn()
'    End Sub
'End Class
