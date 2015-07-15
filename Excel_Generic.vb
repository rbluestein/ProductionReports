Imports System.Runtime.InteropServices.Marshal

Public Class Excel_Generic
    Inherits ExcelBase

#Region " Declarations "
    Private cEnviro As Enviro
    Private cCommon As New Common
    Private cReport As New Report
    Private cAppPath As String
    Private cTemplateFullPath As String
    Private cOutputFileName As String
    Private cOutputFullPath As String
    Private cReportConfig As ReportConfig

    ' ___ Declare the Excel objects
    Private oExcel As Microsoft.Office.Interop.Excel.Application
    Private oBooks As Microsoft.Office.Interop.Excel.Workbooks
    Private oBook As Microsoft.Office.Interop.Excel.Workbook
    Private oSheets As Microsoft.Office.Interop.Excel.Sheets
    Private oSheet As Microsoft.Office.Interop.Excel.Worksheet
    Private oCells As Microsoft.Office.Interop.Excel.Range
#End Region

    Public Enum ExportSourceEnum
        Field = 1
        Collection = 2
        Table = 3
    End Enum

    'Public Function GetRangeData(ByVal RangeName As String)
    '    Dim Range As Microsoft.Office.Interop.Excel.Range
    '    Dim CurDate As String
    '    Dim i As Integer
    '    Dim Coll As Collection

    '    Coll = GetRangeData(RangeName)



    '    oSheet = CType(oSheets.Item(3), Microsoft.Office.Interop.Excel.Worksheet)
    '    For i = 1 To oSheet.Range("Dogs").Count
    '        CurDate = oSheet.Range("Dogs").Cells(i).Value
    '    Next

    'End Function

    Public ReadOnly Property OutputFullPath() As String
        Get
            Return cOutputFullPath
        End Get
    End Property

    Public Function SetFill(ByVal SheetNum As Integer, ByVal RowNum As Integer) As Microsoft.Office.Interop.Excel.Application
        Try
            oSheet = CType(oSheets.Item(SheetNum), Microsoft.Office.Interop.Excel.Worksheet)
            With oSheet.Range(RowNum & ":" & RowNum).Interior
                .ColorIndex = 6
                ' .Pattern = xlSolid
            End With

            Return oExcel

        Catch ex As Exception
            cReport.Report("Error #555.1 : Excel_Generic.SetFill " & ex.Message, Report.ReportTypeEnum.LogError)
        End Try
    End Function

    Public Function DateLookup(ByVal RangeName As String, ByVal Value As DateTime) As Collection
        Dim i As Integer
        Dim Coll As Collection
        Dim Results As New Collection
        Dim CurDate As Date
        Dim RowOffset As Integer

        Try

            Coll = GetRangeData(RangeName, oBook)
            oSheet = CType(oSheets.Item(Coll("SheetNum")), Microsoft.Office.Interop.Excel.Worksheet)
            For i = 0 To oSheet.Range(RangeName).Count - 1
                CurDate = oSheet.Range(RangeName).Cells(i + 1).Value
                If Date.Compare(CurDate, Value) = 0 Then
                    RowOffset = i
                    Exit For
                End If
            Next
            Results.Add(Coll("SheetNum"), "SheetNum")
            Results.Add(RowOffset, "RowOffset")
            Results.Add(oExcel, "ActiveExcel")
            Return Results
        Catch ex As Exception
            ' Throw New Exception("Error #557: ExcelOut GetRangeData " & ex.Message, ex)
            cReport.Report("Excel_Generic.DateLookup  #555.2 " & ex.Message, Report.ReportTypeEnum.LogError)
        End Try
    End Function

    Public Function GetDateRow(ByVal RangeName As String, ByVal Value As DateTime) As Integer
        Dim i As Integer
        Dim Coll As Collection
        Dim CurDate As Date
        Dim RowOffset As Integer
        Dim BaseRow As Integer


        Try

            Coll = GetRangeData(RangeName, oBook)
            BaseRow = Coll("xlAddr1RowNum")
            oSheet = CType(oSheets.Item(Coll("SheetNum")), Microsoft.Office.Interop.Excel.Worksheet)
            For i = 0 To oSheet.Range(RangeName).Count - 1
                CurDate = oSheet.Range(RangeName).Cells(i + 1).Value
                If Date.Compare(CurDate, Value) = 0 Then
                    RowOffset = i
                    Exit For
                End If
            Next
            Return BaseRow + RowOffset
        Catch ex As Exception
            ' Throw New Exception("Error #557: ExcelOut GetRangeData " & ex.Message, ex)
            cReport.Report("Excel_Generic.GetDateRow  #555.3 " & ex.Message, Report.ReportTypeEnum.LogError)
        End Try
    End Function

    Public Sub New(ByVal ReportConfig As ReportConfig)
        Try
            cReportConfig = ReportConfig
            cEnviro = gEnviro
            oExcel = New Microsoft.Office.Interop.Excel.Application

            ' ___ Identify the template and the output files 
            cAppPath = cEnviro.GetAppPath
            cTemplateFullPath = cAppPath & "\Templates\" & cReportConfig.TemplateName
            'cOutputFileName = "ProdRpts_" & cReportConfig.ReportName & "_" & cCommon.GetServerDateTime.ToString("yyyyMMdd_HHmmss") & ".xls"
            cOutputFileName = "ProdRpts_" & ReportConfig.ReportName & "_" & cEnviro.ReportDateTime & ".xls"
            cOutputFullPath = cAppPath & "\TempData\" & cOutputFileName

            ' ___ Configure Excel
            oExcel.Visible = False
            oExcel.DisplayAlerts = False

            '___ Start a new workbook 
            oBooks = oExcel.Workbooks
            oBooks.Open(cTemplateFullPath)
            oBook = oBooks.Item(1)
            oSheets = oBook.Worksheets

        Catch ex As Exception
            ' Throw New Exception("Error #557: ExcelOut GetRangeData " & ex.Message, ex)
            cReport.Report("Excel_Generic.New  #555.4 " & ex.Message, Report.ReportTypeEnum.LogError)
        End Try
    End Sub

    Public Sub Finish(ByRef ActiveExcel As Microsoft.Office.Interop.Excel.Application)

        Try
            oExcel = ActiveExcel

            ' ___ Save and close the workbook
            oBook.SaveAs(cOutputFullPath)
            oBook.Close()

            ' ___ Quit Excel and thoroughly deallocate everything
            oExcel.Quit()
            ReleaseComObject(oCells)
            ReleaseComObject(oSheet)
            ReleaseComObject(oSheets)
            ReleaseComObject(oBook)
            ReleaseComObject(oBooks)
            ReleaseComObject(oExcel)
            oExcel = Nothing
            oBooks = Nothing
            oBook = Nothing
            oSheets = Nothing
            oSheet = Nothing
            oCells = Nothing
            System.GC.Collect()
        Catch ex As Exception
            cReport.Report("Excel_Generic.Finish  #555.5 " & ex.Message, Report.ReportTypeEnum.LogError)
        End Try
    End Sub

    Public Function ExportFieldToExcelForSum(ByVal Value As String, ByRef ExcelAddress As ExcelAddress, Optional ByRef ActiveExcel As Microsoft.Office.Interop.Excel.Application = Nothing) As Results
        Dim Results As Results
        Results = ExportToExcel(ExportSourceEnum.Field, Value, ExcelAddress, True, ActiveExcel)
        Return Results
    End Function

    Public Function ExportFieldToExcel(ByVal Value As String, ByRef ExcelAddress As ExcelAddress, Optional ByRef ActiveExcel As Microsoft.Office.Interop.Excel.Application = Nothing) As Results
        Dim Results As Results
        Results = ExportToExcel(ExportSourceEnum.Field, Value, ExcelAddress, False, ActiveExcel)
        Return Results
    End Function

    Public Function ExportCollectionToExcel(ByVal Coll As Collection, ByRef ExcelAddress As ExcelAddress, Optional ByRef ActiveExcel As Microsoft.Office.Interop.Excel.Application = Nothing) As Results
        Dim Results As Results
        Results = ExportToExcel(ExportSourceEnum.Collection, Coll, ExcelAddress, False, ActiveExcel)
        Return Results
    End Function

    Public Function ExportTableToExcel(ByVal dt As DataTable, ByRef ExcelAddress As ExcelAddress, Optional ByRef ActiveExcel As Microsoft.Office.Interop.Excel.Application = Nothing) As Results
        Dim Results As Results
        Results = ExportToExcel(ExportSourceEnum.Table, dt, ExcelAddress, False, ActiveExcel)
        Return Results
    End Function

    Public Function ExportToExcel(ByVal ExportSource As ExportSourceEnum, ByVal Value As Object, ByRef ExcelAddress As ExcelAddress, ByVal ForSum As Boolean, ByRef ActiveExcel As Microsoft.Office.Interop.Excel.Application) As Results
        Dim i As Integer
        Dim dtRowNum As Integer
        Dim MyResults As New Results
        Dim RangeDataColl As Collection
        Dim Common As New Common
        Dim Coll As Collection
        Dim dt As DataTable
        Dim StartRow As Integer
        Dim StartColumn As Integer
        Dim SheetNum As Integer
        Dim dr As DataRow
        Dim ItemArray() As Object
        Dim ErrorInd As Boolean

        Try

            If Not ActiveExcel Is Nothing Then
                oExcel = ActiveExcel
            End If

            Select Case ExportSource
                Case ExportSourceEnum.Collection
                    Coll = Value
                Case ExportSourceEnum.Table
                    dt = Value
            End Select

            If ExcelAddress.RangeName = Nothing Then
                StartRow = ExcelAddress.RowNum
                StartColumn = ExcelColumnToNumber(ExcelAddress.ColumnLtr)
                SheetNum = ExcelAddress.SheetNum
            Else
                RangeDataColl = GetRangeData(ExcelAddress.RangeName, oBook)
                StartRow = RangeDataColl("xlAddr1RowNum")
                StartColumn = ExcelColumnToNumber(RangeDataColl("xlAddr1ColumnName"))
                SheetNum = RangeDataColl("SheetNum")
            End If

            ' ___ Set zoom
            oExcel.ActiveWindow.Zoom = 100

            ' ___ Identify the worksheet
            oSheet = CType(oSheets.Item(SheetNum), Microsoft.Office.Interop.Excel.Worksheet)

            ' ___ Create a range from the the worksheet
            oCells = oSheet.Cells

            ' ___ Start cell
            StartRow += ExcelAddress.RowOffset
            StartColumn += ExcelAddress.ColumnOffset

            Select Case ExportSource
                Case ExportSourceEnum.Field
                    If ForSum Then
                        oCells(StartRow, StartColumn) = oCells(StartRow, StartColumn).value + CType(Value, System.Decimal)
                        oCells(StartRow, StartColumn).Interior.ColorIndex = 15
                    Else
                        oCells(StartRow, StartColumn) = Value
                    End If


                Case ExportSourceEnum.Collection
                    For i = 1 To Coll.Count
                        oCells(StartRow, StartColumn) = Coll(i)
                        StartColumn += 1
                    Next

                Case ExportSourceEnum.Table
                    For dtRowNum = 0 To dt.Rows.Count - 1
                        dr = dt.Rows.Item(dtRowNum)
                        ItemArray = dr.ItemArray
                        For i = 0 To ItemArray.GetUpperBound(0)
                            oCells(StartRow + dtRowNum, StartColumn + i) = ItemArray(i).ToString
                        Next
                    Next
            End Select

            MyResults.Success = True
            MyResults.Value = "./TempData/" & cOutputFileName
            MyResults.Value = oExcel
            Return MyResults

            '"./TempData/CCM_20090519_204900.xls"

        Catch ex As Exception
            MyResults.Success = False
            MyResults.Message = "Error #554.2: Excel ExportToExcel " & ex.Message
            Return MyResults
        Finally
            If ErrorInd Then
                Try
                    Finish(oExcel)
                Catch ex As Exception
                    cReport.Report("Excel_Generic.ExportToExcel  #555.6 " & ex.Message, Report.ReportTypeEnum.LogError)
                End Try
            End If
        End Try
    End Function

    'Public Function ExportFieldToExcel(ByVal Value As String, ByVal SheetNum As Integer, ByVal RangeName As String, ByVal RowOffset As Integer, Optional ByRef ActiveExcel As Microsoft.Office.Interop.Excel.Application = Nothing) As Results
    '    Dim MyResults As New Results
    '    Dim RangeDataColl As Collection
    '    Dim Common As New Common
    '    Dim AppPath As String
    '    Dim LoadDataResults As Results
    '    Dim Column As Integer
    '    Dim Row As Integer

    '    Try

    '        If Not ActiveExcel Is Nothing Then
    '            oExcel = ActiveExcel
    '        End If

    '        ' ___ Get the anchor range data for this segment
    '        RangeDataColl = GetRangeData(RangeName, oBook)

    '        ' ___ Identify the worksheet
    '        oSheet = CType(oSheets.Item(SheetNum), Microsoft.Office.Interop.Excel.Worksheet)

    '        ' ___ Create a range from the the worksheet
    '        oCells = oSheet.Cells

    '        ' ___ Cell address
    '        Column = ExcelColumnToNumber(RangeDataColl("xlAddr1ColumnName"))
    '        Row = RangeDataColl("xlAddr1RowNum") + RowOffset


    '        ' ___ Populate the cell with the value
    '        oCells(Row + RowOffset, Column) = Value

    '        ' ___ Output

    '        MyResults.Success = True
    '        MyResults.Value = "./TempData/" & cOutputFileName
    '        MyResults.Value = oExcel
    '        Return MyResults

    '        '"./TempData/CCM_20090519_204900.xls"

    '    Catch ex As Exception
    '        MyResults.Success = False
    '        MyResults.Message = "Error #555: Excel ExportFieldToExcel " & ex.Message
    '        Return MyResults
    '    End Try
    'End Function

    'Public Function ExportCollectionToExcel(ByRef Coll As Collection, ByVal SheetNum As Integer, ByVal RangeName As String, ByVal RowOffset As Integer, Optional ByRef ActiveExcel As Microsoft.Office.Interop.Excel.Application = Nothing) As Results
    '    Dim i As Integer
    '    Dim MyResults As New Results
    '    Dim RangeDataColl As Collection
    '    Dim Common As New Common
    '    Dim AppPath As String
    '    Dim LoadDataResults As Results
    '    Dim Column As Integer
    '    Dim Row As Integer

    '    Try

    '        If Not ActiveExcel Is Nothing Then
    '            oExcel = ActiveExcel
    '        End If

    '        ' ___ Get the anchor range data for this segment
    '        RangeDataColl = GetRangeData(RangeName, oBook)

    '        ' ___ Identify the worksheet
    '        oSheet = CType(oSheets.Item(SheetNum), Microsoft.Office.Interop.Excel.Worksheet)

    '        ' ___ Create a range from the the worksheet
    '        oCells = oSheet.Cells

    '        ' ___ Cell address
    '        Column = ExcelColumnToNumber(RangeDataColl("xlAddr1ColumnName"))
    '        Row = RangeDataColl("xlAddr1RowNum")

    '        ' ___ Populate the spreadsheet with the collection values
    '        For i = 1 To Coll.Count
    '            oCells(Row + RowOffset, Column) = Coll(i)
    '            Column += 1
    '        Next

    '        ' ___ Output
    '        MyResults.Success = True
    '        MyResults.Value = "./TempData/" & cOutputFileName
    '        MyResults.Value = oExcel
    '        Return MyResults

    '        '"./TempData/CCM_20090519_204900.xls"

    '    Catch ex As Exception
    '        MyResults.Success = False
    '        MyResults.Message = "Error #555: Excel ExportFieldToExcel " & ex.Message
    '        Return MyResults
    '    End Try
    'End Function

    'Public Function ExportTableToExcel(ByRef dt As DataTable, ByVal SheetNum As Integer, ByVal RangeName As String, ByVal RowOffset As Integer, Optional ByRef ActiveExcel As Microsoft.Office.Interop.Excel.Application = Nothing) As Results
    '    Dim MyResults As New Results
    '    Dim RangeDataColl As Collection
    '    Dim Common As New Common
    '    Dim AppPath As String
    '    Dim LoadDataResults As Results

    '    Try

    '        If Not ActiveExcel Is Nothing Then
    '            oExcel = ActiveExcel
    '        End If


    '        ' ___ Get the anchor range data for this segment
    '        RangeDataColl = GetRangeData(RangeName, oBook)

    '        ' ___ Identify the worksheet
    '        'oSheet = CType(oSheets.Item(RangeDataColl("SheetNum")), Microsoft.Office.Interop.Excel.Worksheet)
    '        oSheet = CType(oSheets.Item(SheetNum), Microsoft.Office.Interop.Excel.Worksheet)

    '        ' ___ Create a range from the the worksheet
    '        oCells = oSheet.Cells

    '        ' ____ Populate the worksheet with the segment data
    '        LoadDataResults = LoadData(dt, RangeDataColl, RowOffset, oCells, oBook)
    '        If Not LoadDataResults.Success Then
    '            MyResults.Success = False
    '            MyResults.Message = LoadDataResults.Message
    '            Return MyResults
    '        End If

    '        ' ___ Output
    '        System.Diagnostics.Debug.WriteLine("Excel.ExportToExcel : Processing segment " & RangeName)

    '        MyResults.Success = True
    '        MyResults.Value = "./TempData/" & cOutputFileName
    '        MyResults.Value = oExcel
    '        Return MyResults

    '        '"./TempData/CCM_20090519_204900.xls"

    '    Catch ex As Exception
    '        MyResults.Success = False
    '        MyResults.Message = "Error #554: Excel TableExportToExcel " & ex.Message
    '        Return MyResults
    '    End Try
    'End Function

    'Private Function LoadData(ByRef dt As DataTable, ByRef RangeDataColl As Collection, ByVal RowOffset As Integer, ByRef oCells As Microsoft.Office.Interop.Excel.Range, ByRef oBook As Microsoft.Office.Interop.Excel.Workbook) As Results
    '    Dim i As Integer
    '    Dim dr As DataRow
    '    Dim ItemArray() As Object
    '    Dim dtRowNum As Integer
    '    Dim StartRow As Integer
    '    Dim StartColumn As Integer
    '    Dim MyResults As New Results

    '    Try

    '        ' ___ Get the start column and row
    '        StartColumn = ExcelColumnToNumber(RangeDataColl("xlAddr1ColumnName"))
    '        StartRow = RangeDataColl("xlAddr1RowNum") + RowOffset

    '        ' ___ Output Data
    '        For dtRowNum = 0 To dt.Rows.Count - 1
    '            dr = dt.Rows.Item(dtRowNum)
    '            ItemArray = dr.ItemArray
    '            For i = 0 To ItemArray.GetUpperBound(0)
    '                oCells(StartRow + dtRowNum, StartColumn + i) = ItemArray(i).ToString
    '            Next
    '        Next

    '        MyResults.Success = True
    '        Return MyResults

    '    Catch ex As Exception
    '        MyResults.Success = False
    '        MyResults.Message = "Error #555b Excel.LoadData " & ex.Message
    '        Return MyResults
    '    End Try

    'End Function

    Public Function GetRangeData(ByVal RangeName As String) As Collection
        Dim Coll As Collection
        Coll = GetRangeData(RangeName, oBook)
        Return Coll
    End Function

    Private Function GetRangeData(ByVal RangeName As String, ByRef oBook As Microsoft.Office.Interop.Excel.Workbook) As Collection
        Dim i As Integer
        Dim xlRangeName As String
        Dim SearchRangeName As String
        Dim Address As Object
        Dim Box1() As String
        Dim SheetName As String
        Dim Coll As New Collection

        Try

            SearchRangeName = Replace(RangeName, "|", "")
            For i = 1 To oBook.Names.Count
                xlRangeName = oBook.Names.Item(i).Name
                If xlRangeName = SearchRangeName Then
                    Address = oBook.Names.Item(i)._Default
                    Box1 = Split(Address, "!")
                    SheetName = Box1(0).Substring(1)
                    SheetName = Replace(SheetName, "'", "")
                    'Box2 = Split(SheetName, "Sheet")
                    'SheetNum = Box2(1)
                    'Box1 = Split(Box1(1), "$")

                    Box1 = Split(Box1(1), "$")

                    'xlAddr1ColumnName = Box1(1)
                    'xlAddr1RowNum = Box1(2)

                    Coll.Add(SheetName, "SheetName")
                    Coll.Add(Replace(Box1(2), ":", ""), "xlAddr1RowNum")
                    Coll.Add(Box1(1), "xlAddr1ColumnName")

                    If Box1.GetUpperBound(0) = 4 Then
                        Coll.Add(Box1(4), "xlAddr2RowNum")
                        Coll.Add(Box1(3), "xlAddr2ColumnName")
                    End If


                    Dim Sheet As Microsoft.Office.Interop.Excel.Worksheet
                    For Each Sheet In oBook.Worksheets
                        If Sheet.Name.ToLower = SheetName.ToLower Then
                            Coll.Add(Sheet.Index, "SheetNum")
                            Coll.Add(Sheet.Range(xlRangeName).Cells(1, 1).Value, "FirstValue")
                            Exit For
                        End If

                    Next

                    Coll.Add(oExcel, "ActiveExcel")
                    Return Coll
                End If
            Next

            ' ___ Range not found
            cReport.Report("Excel_Generic.GetRangeData  #555.7a Unable to find range " & SearchRangeName, Report.ReportTypeEnum.LogError)


        Catch ex As Exception
            ' Throw New Exception("Error #557: ExcelOut GetRangeData " & ex.Message, ex)
            cReport.Report("Excel_Generic.GetRangeData  #555.7 RangeName: " & RangeName & "   " & ex.Message, Report.ReportTypeEnum.LogError)
        End Try
    End Function
End Class
