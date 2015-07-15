Imports Microsoft.VisualBasic
Imports System.Data

Imports System.Data.SqlClient
Imports Microsoft.Office.Interop.Excel

Public Class exceldt


    ' Com -> Microsoft Excel 11.0 Object library

    '' Start a new workbook in Excel.
    '' ===========================
    ''oExcel = CreateObject("Excel.Application")
    'Dim oExcel As New Microsoft.Office.Interop.Excel.Application
    '            oBook = oExcel.Workbooks.Add
    ''oSheet = oBook.Worksheets(1)
    'Dim oSheet As New Microsoft.Office.Interop.Excel.Worksheet
    '            oSheet = oBook.Worksheets(1)
    '            oSheet.Range("1:1").Font.Bold = True

    Public Shared Sub ExportToExcel(ByRef Coll As CollX)
        Dim i As Integer
        Dim oBook As Object
        Dim SheetColl As New CollX

        Try

            ' ___ Start a new workbook in Excel.
            Dim oExcel As New Microsoft.Office.Interop.Excel.Application
            oBook = oExcel.Workbooks.Add
            'oSheet = oBook.Worksheets(1)

            For i = 1 To Coll.Count
                ExportToExcel2(oExcel, oBook, Coll, i, SheetColl)
            Next

            'oBook.Worksheets("Call Incomplete").Select()
            oBook.Worksheets(1).Select()

            oExcel.Visible = True
            oExcel.DisplayAlerts = False

            For i = 1 To SheetColl.Count
                Dim oSheet As Object
                oSheet = SheetColl(i)
                oSheet = Nothing
            Next

            'oSheet = Nothing
            oBook = Nothing
            'oExcel.Quit()
            oExcel = Nothing
            GC.Collect()

        Catch ex As Exception
            Throw New Exception("Error #2401: Excel.ExportToExcel (Collx). " & ex.Message)
        End Try
    End Sub

    Private Shared Sub ExportToExcel2(ByRef oExcel As Microsoft.Office.Interop.Excel.Application, ByRef oBook As Object, ByRef Coll As CollX, ByVal Num As Integer, ByRef SheetColl As CollX)
        Dim i As Integer
        Dim DataArray As Object
        Dim dt As System.Data.DataTable
        Dim ExcelRow As Integer
        Dim ExcelCol As String
        Dim DataType As String
        Dim CharCount As Integer
        Dim Max As Integer = 911 '2000 '911
        Dim NumberDimensions As Integer
        Dim TooBig As New Collection
        Dim Row As Integer
        Dim Col As Integer

        Try

            dt = Coll(Num)

            Dim oSheet As New Microsoft.Office.Interop.Excel.Worksheet

            If Num > 3 Then
                oBook.Worksheets.Add()
                'oBook.Worksheets.Remove()
                Dim SheetName As String
                SheetName = "Sheet" & Num
                oBook.worksheets(SheetName).Move(After:=oBook.worksheets(Num))
            End If
            oSheet = oBook.Worksheets(Num)
            SheetColl.Assign(Num, oSheet)
            oSheet.Range("1:1").Font.Bold = True

            ReDim DataArray(dt.Rows.Count, dt.Columns.Count - 1)

            ' ___ Column headings
            For Col = 0 To dt.Columns.Count - 1
                DataArray(0, Col) = dt.Columns(Col).ColumnName
            Next
            ExcelRow = 2
            ExcelCol = "A"

            ' ___ Data
            For Row = 0 To dt.Rows.Count - 1
                For Col = 0 To dt.Columns.Count - 1

                    DataType = dt.Columns.Item(Col).DataType.ToString().ToUpper

                    If DataType = "SYSTEM.GUID" Then
                        If IsDBNull(dt.Rows(Row)(Col)) Then
                            DataArray(Row + 1, Col) = String.Empty
                        Else
                            DataArray(Row + 1, Col) = dt.Rows(Row)(Col).ToString
                        End If
                    ElseIf DataType = "SYSTEM.STRING" Then
                        If IsDBNull(dt.Rows(Row)(Col)) Then
                            DataArray(Row + 1, Col) = String.Empty
                        Else
                            CharCount = dt.Rows(Row)(Col).length
                            If CharCount > Max Then
                                Dim BigData(1) As Object
                                BigData(0) = ExcelCol & ExcelRow.ToString
                                BigData(1) = dt.Rows(Row)(Col)
                                TooBig.Add(BigData)
                                DataArray(Row + 1, Col) = String.Empty
                            Else
                                DataArray(Row + 1, Col) = dt.Rows(Row)(Col)
                            End If
                        End If
                    Else
                        If IsDBNull(dt.Rows(Row)(Col)) Then
                            DataArray(Row + 1, Col) = String.Empty
                        Else
                            DataArray(Row + 1, Col) = dt.Rows(Row)(Col)
                        End If
                    End If
                    ExcelCol = AddColumn(ExcelCol)
                Next
                ExcelCol = "A"
                ExcelRow += 1
            Next
            If dt.Columns.Count = 1 Then
                NumberDimensions = 1
            End If

            If NumberDimensions = 1 Then
                Dim TempArray(DataArray.getupperbound(0), 1) As Object
                For i = 0 To TempArray.GetUpperBound(0)
                    TempArray(i, 0) = DataArray(i, 0)
                    TempArray(i, 1) = String.Empty
                Next
                DataArray = TempArray
            End If

            oSheet.Range("A1").Resize(DataArray.GetUpperBound(0) + 1, DataArray.GetUpperBound(1) + 1).Value = DataArray
            If TooBig.Count > 0 Then
                For i = 1 To TooBig.Count
                    oSheet.Range(TooBig(i)(0)).Value = TooBig(i)(1)
                Next
            End If

            If (Not IsNumeric(Coll.Key(Num))) Then
                oSheet.Name = Coll.Key(Num)
            End If
            oSheet.Columns.AutoFit()
            DataArray = Nothing

        Catch ex As Exception
            Throw New Exception("Error #2402: Excel.ExportToExcel (Collx). " & ex.Message)
        End Try
    End Sub

    Public Shared Sub ee(ByRef dt As System.Data.DataTable)
        ExportToExcel(dt)
    End Sub

    Public Shared Sub ExportToExcel(ByRef dt As System.Data.DataTable, Optional ByVal SheetName As String = "")
        Dim i As Integer
        Dim oBook As Object
        Dim Row As Integer
        Dim Col As Integer
        Dim DataArray As Object
        Dim CharCount As Integer
        Dim DataType As String
        Dim Max As Integer = 911 '2000 '911
        Dim NumberDimensions As Integer
        Dim TooBig As New Collection
        Dim ExcelRow As Integer
        Dim ExcelCol As String

        Try

            ' ___ Start a new workbook in Excel.
            Dim oExcel As New Microsoft.Office.Interop.Excel.Application
            oBook = oExcel.Workbooks.Add
            'oSheet = oBook.Worksheets(1)
            Dim oSheet As New Microsoft.Office.Interop.Excel.Worksheet
            oSheet = oBook.Worksheets(1)
            oSheet.Range("1:1").Font.Bold = True
            oSheet.Name = "Blubbus"


            ReDim DataArray(dt.Rows.Count, dt.Columns.Count - 1)

            ' ___ Column headings
            For Col = 0 To dt.Columns.Count - 1
                DataArray(0, Col) = dt.Columns(Col).ColumnName
            Next
            ExcelRow = 2
            ExcelCol = "A"

            ' ___ Data
            For Row = 0 To dt.Rows.Count - 1
                For Col = 0 To dt.Columns.Count - 1

                    DataType = dt.Columns.Item(Col).DataType.ToString().ToUpper

                    If DataType = "SYSTEM.GUID" Then
                        If IsDBNull(dt.Rows(Row)(Col)) Then
                            DataArray(Row + 1, Col) = String.Empty
                        Else
                            DataArray(Row + 1, Col) = dt.Rows(Row)(Col).ToString
                        End If
                    ElseIf DataType = "SYSTEM.STRING" Then
                        If IsDBNull(dt.Rows(Row)(Col)) Then
                            DataArray(Row + 1, Col) = String.Empty
                        Else
                            CharCount = dt.Rows(Row)(Col).length
                            If CharCount > Max Then
                                Dim BigData(1) As Object
                                BigData(0) = ExcelCol & ExcelRow.ToString
                                BigData(1) = dt.Rows(Row)(Col)
                                TooBig.Add(BigData)
                                DataArray(Row + 1, Col) = String.Empty
                            Else
                                DataArray(Row + 1, Col) = dt.Rows(Row)(Col)
                            End If
                        End If
                    Else
                        If IsDBNull(dt.Rows(Row)(Col)) Then
                            DataArray(Row + 1, Col) = String.Empty
                        Else
                            DataArray(Row + 1, Col) = dt.Rows(Row)(Col)
                        End If
                    End If
                    ExcelCol = AddColumn(ExcelCol)
                Next
                ExcelCol = "A"
                ExcelRow += 1
            Next
            If dt.Columns.Count = 1 Then
                NumberDimensions = 1
            End If

            If NumberDimensions = 1 Then
                Dim TempArray(DataArray.getupperbound(0), 1) As Object
                For i = 0 To TempArray.GetUpperBound(0)
                    TempArray(i, 0) = DataArray(i, 0)
                    TempArray(i, 1) = String.Empty
                Next
                DataArray = TempArray
            End If

            oSheet.Range("A1").Resize(DataArray.GetUpperBound(0) + 1, DataArray.GetUpperBound(1) + 1).Value = DataArray
            If TooBig.Count > 0 Then
                For i = 1 To TooBig.Count
                    oSheet.Range(TooBig(i)(0)).Value = TooBig(i)(1)
                Next
            End If

            If SheetName <> Nothing Then
                If SheetName.Length > 0 Then
                    oSheet.Name = SheetName
                End If
            End If
            oSheet.Columns.AutoFit()

            DataArray = Nothing
            oExcel.Visible = True
            oExcel.DisplayAlerts = False
            oSheet = Nothing
            oBook = Nothing
            'oExcel.Quit()
            oExcel = Nothing
            GC.Collect()

        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine(ex.Message)
        End Try
    End Sub


    Public Shared Sub IndividualCellExportToExcel(ByVal dt As System.Data.DataTable)
        Dim oExcel As Object
        Dim oBook As Object
        Dim oSheet As Object
        Dim Cols As Integer
        Dim Address As String
        Dim ColName As String = "A"
        Dim PrintRow As Integer

        'Start a new workbook in Excel.
        oExcel = CreateObject("Excel.Application")
        oBook = oExcel.Workbooks.Add

        'Add data to cells of the first worksheet in the new workbook.
        oSheet = oBook.Worksheets(1)
        'oSheet.Range("A1:B1").Font.Bold = True
        oSheet.Range("1:1").Font.Bold = True

        If dt.Rows.Count > 65536 Then
            'Dim errorobj As New ErrorClass
            'errorobj.SimpleHandleError(EnumClass.ErrorSourceUserAppIOItem.UserError, "Number of rows exceeds Excel's maximum.", True)
        End If

        'Heading
        PrintRow = 1
        ColName = "A"
        For Cols = 0 To dt.Columns.Count - 1
            Address = ColName & PrintRow.ToString
            oSheet.range(Address).value = dt.Columns(Cols).ColumnName
            ColName = AddColumn(ColName)
        Next

        ' Data
        For PrintRow = 2 To dt.Rows.Count + 1
            ColName = "A"
            For Cols = 0 To dt.Columns.Count - 1
                Address = ColName & PrintRow.ToString
                oSheet.range(Address).value = dt.Rows(PrintRow - 2)(Cols)
                ColName = AddColumn(ColName)
            Next
        Next

        oExcel.Visible = True
        oExcel.DisplayAlerts = False
        oSheet = Nothing
        oBook = Nothing
        oExcel = Nothing
        GC.Collect()
    End Sub

    Public Shared Function AddColumn(ByVal ColName As String) As String
        Dim i As Integer
        Dim MaxLength As Integer = 5
        Dim Rec_(MaxLength - 1) As String
        Dim Results As String = String.Empty
        Dim Carry As Boolean

        ' Pad ColName
        ColName = ColName.PadLeft(MaxLength)

        ' Load the column value into Rec_
        For i = 0 To MaxLength - 1
            Rec_(i) = ColName.Substring(i, 1)
        Next

        ' Begin processing
        For i = MaxLength - 1 To 0 Step -1
            If i = MaxLength - 1 Then
                ' Perform alpha addition on the rightmost character.Apply carry if applicable.
                If Rec_(i) = "Z" Then
                    Rec_(i) = "A"
                    Carry = True
                Else
                    Rec_(i) = Chr(Asc(Rec_(i)) + 1)
                End If
            Else   ' Not the rightmost character

                ' If the character to its right resulted in a carry...
                If Carry Then
                    If Asc(Rec_(i)) = 32 Then
                        Rec_(i) = "A"
                        Carry = False
                    ElseIf Rec_(i) = "Z" Then
                        Rec_(i) = "A"
                        Carry = True
                    Else
                        Rec_(i) = Chr(Asc(Rec_(i)) + 1)
                        Carry = False
                    End If

                End If
            End If
        Next

        For i = 0 To Rec_.GetUpperBound(0)
            Results &= Rec_(i)
        Next
        Return Trim(Results)
    End Function


    Public Function ExcelToDT(ByVal FullPath As String, ByVal Header As Boolean, Optional ByVal Sql As String = "") As System.Data.DataTable
        'Dim ConnStr As String
        'ConnStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Inetpub\wwwroot\ImportExportManagement\MasterList.xls;Extended Properties=Excel 8.0;"
        Dim dt As System.Data.DataTable = Nothing

        Try

            dt = New System.Data.DataTable
            If Sql.Length = 0 Then
                Sql = "SELECT * FROM [Sheet1$]"
            End If

            'Dim da As New OleDbDataAdapter("SELECT * FROM [Feed_DTS$]", ConnStr)
            '   Dim da As New System.Data.OleDb.OleDbDataAdapter("SELECT * FROM [ApptLicFieldsOnly$]", GetExcelConnectionString("License.xls", True))
            ' Dim da As New System.Data.OleDb.OleDbDataAdapter("SELECT * FROM [" & WorksheetName & "$]", GetExcelConnectionString(Filename, Header))
            Dim da As New System.Data.OleDb.OleDbDataAdapter(Sql, GetExcelConnectionString(FullPath, Header))
            da.Fill(dt)
            Return dt

        Catch ex As Exception
            Stop
            Return Nothing
        End Try

    End Function


    Private Function GetExcelConnectionString(ByVal FullPath As String, ByVal Header As Boolean) As String
        ' http://www.connectionstrings.com/?carrier=excel


        Dim ConnString As String
        ConnString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=<fullpath>;Extended Properties=""Excel 8.0;HDR=<header>"";"

        If System.Environment.MachineName.ToUpper = "LT-5ZFYRC1" Then
            ConnString = Replace(ConnString, "<fullpath>", FullPath)
            If Header Then
                ConnString = Replace(ConnString, "<header>", "Yes")
            Else
                ConnString = Replace(ConnString, "<header>", "No")
            End If
            Return ConnString

        Else
            ' Return "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Inetpub\wwwroot\UserManagement\ProdBVIUSERSUSERScEnrollerStatusA.xls;Extended Properties=Excel 8.0;"
            Return Nothing
        End If
    End Function

    Public Function ExcelToTbl(ByVal FullPath As String, Optional ByVal Sql As String = "") As System.Data.DataTable
        Dim DataAdapter As SqlDataAdapter
        Dim dt As New System.Data.DataTable
        Dim ConnString As String

        If Sql.Length = 0 Then
            Sql = "SELECT * FROM [Sheet1$]"
        End If
        'ConnString = "provider=Microsoft.Jet.OLEDB.4.0; " & "data source='" & FullPath & " '; " & "Extended Properties=Excel 8.0;"
        ConnString = "data source='" & FullPath & " '; " & "Extended Properties=Excel 8.0;"


        '"user id=BVI_SQL_SERVER;password=noisivtifeneb;database=|;server="

        Dim SqlCmd As New SqlCommand(Sql)
        SqlCmd.CommandType = CommandType.Text
        SqlCmd.Connection = New SqlConnection(ConnString)
        DataAdapter = New SqlDataAdapter(SqlCmd)
        DataAdapter.Fill(dt)
        DataAdapter.Dispose()
        SqlCmd.Dispose()
        Return dt
    End Function


    'Public Function ExcelToTbl()
    '    Dim MyConnection As System.Data.OleDb.OleDbConnection
    '    Dim FullPath As String

    '    Try
    '        FullPath = "C:\Apps\UserManagement\MigrationOutput.xls"
    '        ''''''' Fetch Data from Excel
    '        Dim DtSet As System.Data.DataSet
    '        Dim MyCommand As System.Data.OleDb.OleDbDataAdapter
    '        MyConnection = New System.Data.OleDb.OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; " & "data source='" & FullPath & " '; " & "Extended Properties=Excel 8.0;")

    '        ' Select the data from Sheet1 of the workbook.
    '        MyCommand = New System.Data.OleDb.OleDbDataAdapter("select * from [Sheet1$]", MyConnection)
    '        MyCommand.TableMappings.Add("Table", "Attendence")
    '        DtSet = New System.Data.DataSet
    '        MyCommand.Fill(DtSet)

    '        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '        'DataGrid1.DataSource = DtSet.Tables(0)
    '        MyConnection.Close()

    '    Catch ex As Exception
    '        MyConnection.Close()
    '    End Try
    'End Function

    ' *  Connection String
    '       Syntax: Provider=Microsoft.Jet.OLEDB.4.0;Data Source=<Full Path of Excel File>; Extended Properties="Excel 8.0; HDR=No; IMEX=1".

    'Definition of Extended Properties:
    '    * Excel = <No>
    '      One should specify the version of Excel Sheet here. For Excel 2000 and above, it is set it to Excel 8.0 and for all others, it is Excel 5.0.

    '    * HDR= <Yes/No>
    '      This property will be used to specify the definition of header for each column. If the value is ‘Yes’, the first row will be treated as heading. Otherwise, the heading will be generated by the system like F1, F2 and so on.

    '    * IMEX= <0/1/2>
    '      IMEX refers to IMport EXport mode. This can take three possible values.
    '          o IMEX=0 and IMEX=2 will result in ImportMixedTypes being ignored and the default value of ‘Majority Types’ is used. In this case, it will take the first 8 rows and then the data type for each column will be decided.
    '          o IMEX=1 is the only way to set the value of ImportMixedTypes as Text. Here, everything will be treated as text. 

    'For more info regarding Extended Properties, http://www.dicks-blog.com/archives/2004/06/03/
End Class

