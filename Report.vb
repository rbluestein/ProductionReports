Public Class Report
    Private cEnviro As Enviro
    Private cCommon As Common

    Public Enum ReportTypeEnum
        LogInformation = 1
        LogError = 2
    End Enum
    Public Sub New()
        cEnviro = gEnviro
        cCommon = New Common
    End Sub

    Public Sub Report(ByVal ErrorMessage As String, ByVal ReportType As ReportTypeEnum)
        Dim Coll As New Collection
        Dim SendEmailResults As Results
        Dim EmailSuccess As Boolean
        Dim SendEmailPlease As Boolean
        Dim ShutdownPlease As Boolean
        Dim WriteToLogPlease As Boolean

        Try

            'If Environment.MachineName = "LT-5ZFYRC1" Then
            '    FrmMessage.txtErrorReport.Text = ErrorMessage
            'Else
            '    MessageBox.Show("An error has occurred: " & Environment.NewLine & ErrorMessage, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            'End If

            Select Case ReportType
                Case ReportTypeEnum.LogInformation
                    WriteToLogPlease = True
                    SendEmailPlease = False
                    ShutdownPlease = False
                Case ReportTypeEnum.LogError
                    WriteToLogPlease = True
                    SendEmailPlease = True
                    ShutdownPlease = True
            End Select

            ' ___ Do not send email or make log entries for errors occurring on development machine.
            If Environment.MachineName = "LT-5ZFYRC1" Then
                SendEmailPlease = False
                WriteToLogPlease = True
            End If

            ' ___ Write to log
            If WriteToLogPlease Then
                If ShutdownPlease Then
                    WriteToLogFile(ErrorMessage & " ** ERROR FORCING APPLICATION SHUTDOWN **")
                Else
                    WriteToLogFile(ErrorMessage)
                End If
            End If

            ' ___ Send email
            If SendEmailPlease Then
                Coll.Add(cEnviro.LogFileFullPath)
                SendEmailResults = SendEmail("rbluestein@benefitvision.com", "jkleiman@benefitvision.com", Nothing, "ProductionReports error", ErrorMessage, Coll)
                EmailSuccess = SendEmailResults.Success
            End If

            ' ___ Shut down application
            If ShutdownPlease Then
                cCommon.ExitApplication()
                'Application.Exit()
            End If

        Catch ex As Exception
            Throw New Exception("Error #1502: Report Report. " & ex.Message, ex)
        End Try
    End Sub

    Public Function SendEmail(ByVal SendTo As String, ByVal From As String, ByVal cc As String, ByVal Subject As String, ByVal TextBody As String, Optional ByRef AttachmentColl As Collection = Nothing) As Results
        Dim MyResults As New Results
        Dim i As Integer
        Dim CDOConfig As CDO.Configuration
        Dim iMsg As CDO.Message

        Try

            CDOConfig = New CDO.Configuration
            CDOConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing").Value = 2
            ' CDOConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport").Value = 25
            CDOConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver").Value = "mail.benefitvision.com"
            CDOConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate").Value = 1
            CDOConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendusername").Value = "automail"
            CDOConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendpassword").Value = "$bambam2004#"
            CDOConfig.Fields.Update()

            iMsg = New CDO.Message
            iMsg.To = SendTo
            iMsg.From = From
            iMsg.CC = cc
            iMsg.Subject = Subject

            If Not AttachmentColl Is Nothing Then
                For i = 1 To AttachmentColl.Count
                    iMsg.AddAttachment(AttachmentColl(i))
                Next
            End If

            iMsg.Configuration = CDOConfig
            iMsg.TextBody = TextBody
            'imsg.HTMLBody = htmlbody

            iMsg.Send()

            MyResults.Success = True
            Return MyResults

        Catch ex As Exception
            MyResults.Success = False
            MyResults.Message = "Error #1503: " & ex.Message
            Return MyResults

        Finally
            iMsg.Attachments.DeleteAll()
            CDOConfig = Nothing
            iMsg = Nothing
        End Try
    End Function

    'Public Sub SaveRptUpdateLog()
    '    Dim i As Integer
    '    Dim StreamWriter As System.IO.StreamWriter
    '    Dim Message As New System.Text.StringBuilder
    '    Dim Coll As Collection

    '    Coll = cEnviro.ReportTablesUpdateColl
    '    For i = 1 To Coll.Count
    '        Message.Append(Coll(i) & vbCrLf)
    '    Next

    '    Try
    '        StreamWriter = New System.IO.StreamWriter(cEnviro.ReportTablesUpdateLogFileFullPath, True)
    '        StreamWriter.Write(Message.ToString)
    '    Catch
    '    Finally
    '        Try
    '            StreamWriter.Close()
    '        Catch
    '        End Try
    '    End Try

    'End Sub

    Private Function ReadLogFile() As String
        Dim StreamReader As System.IO.StreamReader
        Dim FileText As String


        Try
            StreamReader = New System.IO.StreamReader(cEnviro.LogFileFullPath)
            FileText = StreamReader.ReadToEnd

            'Do While StreamReader.Peek() >= 0
            '    'Console.WriteLine(StreamReader.ReadLine())
            '    x = StreamReader.ReadLine()
            'Loop
            'StreamReader.Close()
            Return FileText

        Catch ex As Exception
            Throw New Exception("Error #1504: Report ReadLogFile. " & ex.Message, ex)
        Finally
            Try
                StreamReader.Close()
            Catch
            End Try
        End Try
    End Function

    Private Sub WriteToLogFile(ByVal Message As String)
        Dim i As Integer
        'Dim FileInfo As System.IO.FileInfo
        Dim StreamWriter As System.IO.StreamWriter

        Try

            Message = Replace(Message, "~", "")
            'FileInfo = New System.IO.FileInfo(cEnviro.LogFileFullPath)

            Try
                StreamWriter = New System.IO.StreamWriter(cEnviro.LogFileFullPath, True)
            Catch
                Dim procList() As System.Diagnostics.Process = System.Diagnostics.Process.GetProcesses()
                For i = 0 To procList.GetUpperBound(0)
                    If procList(i).ProcessName = "notepad" Then
                        procList(i).Kill()
                    End If
                Next
                StreamWriter = New System.IO.StreamWriter(cEnviro.LogFileFullPath, True)
            End Try

            StreamWriter.Write(GetTimeStamp() & Message & vbCrLf)
        Catch
        Finally
            Try
                StreamWriter.Close()
            Catch
            End Try
        End Try
    End Sub

    Private Function GetTimeStamp() As String
        Return "[" & Date.Now.ToUniversalTime.AddHours(-5).ToString & "] "
    End Function
End Class
