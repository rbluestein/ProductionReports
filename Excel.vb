Imports System.Runtime.InteropServices.Marshal

Public Class Excel
    Inherits ExcelBase

    Public Event NotifyForm(ByRef NotifyFormArgs As NotifyFormArgs)
    Dim e As New NotifyFormArgs(NotifyFormArgs.SourceEnum.Excel)

    Private cEnviro As Enviro
    Private cCommon As New Common
    Private cReport As Report
    Private cOutputFullPath As String

    ' ___ Declare the Excel objects
    Dim oExcel As New Microsoft.Office.Interop.Excel.Application
    Dim oBooks As Microsoft.Office.Interop.Excel.Workbooks
    Dim oBook As Microsoft.Office.Interop.Excel.Workbook
    Dim oSheets As Microsoft.Office.Interop.Excel.Sheets
    Dim oSheet As Microsoft.Office.Interop.Excel.Worksheet
    Dim oCells As Microsoft.Office.Interop.Excel.Range

    Public Enum SegmentType
        Client = 1
        Carrier = 2
    End Enum

    Public Sub New()
        cEnviro = gEnviro
    End Sub

    Public ReadOnly Property OutputFullPath() As String
        Get
            Return cOutputFullPath
        End Get
    End Property

    ' Public Function ExportToExcel(ByVal AppPath As String, ByRef ClientID As String, ByVal SheetNum As Integer, ByRef dt As DataTable, ByVal ShowColumnNames As Boolean) As Results

    Public Function ExportToExcel(ByRef ExcelPack As ExcelPack_RptSupervisor, ByVal ShowColumnNames As Boolean, ByRef ReportConfig As ReportConfig) As Results
        Dim i As Integer
        Dim ExcelPackItem As ExcelPack_RptSupervisor.Item
        Dim MyResults As New Results
        Dim TemplateFullPath As String
        Dim OutputFileName As String
        Dim RangeDataColl As Collection
        Dim Common As New Common
        Dim AppPath As String
        Dim LoadDataResults As Results

        Try

            ' ___ Identify the template and the output files 
            AppPath = cEnviro.GetAppPath
            'TemplateFullPath = AppPath & "\Templates\Template_5.0.xls"
            TemplateFullPath = AppPath & "\Templates\" & ReportConfig.TemplateName
            'OutputFileName = "ProdRpts_" & ReportConfig.ReportName & "_" & Common.GetServerDateTime.ToString("yyyyMMdd_HHmmss") & ".xls"
            OutputFileName = "ProdRpts_" & ReportConfig.ReportName & "_" & cEnviro.ReportDateTime & ".xls"
            cOutputFullPath = AppPath & "\TempData\" & OutputFileName

            ' ___ Configure Excel
            oExcel.Visible = False
            oExcel.DisplayAlerts = False

            '___ Start a new workbook 
            oBooks = oExcel.Workbooks
            oBooks.Open(TemplateFullPath)
            oBook = oBooks.Item(1)
            oSheets = oBook.Worksheets

            For i = 1 To ExcelPack.Coll.Count

                ' ___ Extract the segment from the collection
                ExcelPackItem = ExcelPack.Coll(i)

                ' ___ Get the anchor range data for this segment
                RangeDataColl = GetRangeData("Anchor", ExcelPackItem, oBook)

                ' ___ Identify the worksheet
                oSheet = CType(oSheets.Item(RangeDataColl("SheetNum")), Microsoft.Office.Interop.Excel.Worksheet)

                ' ___ Create a range from the the worksheet
                oCells = oSheet.Cells

                ' ____ Populate the worksheet with the segment data
                LoadDataResults = LoadData(ExcelPackItem, RangeDataColl, oCells, oBook, ShowColumnNames)
                If Not LoadDataResults.Success Then
                    MyResults.Success = False
                    MyResults.Message = LoadDataResults.Message
                    Return MyResults
                End If

                ' ___ Output
                System.Diagnostics.Debug.WriteLine("Excel.ExportToExcel : Processing segment " & i.ToString & " of " & ExcelPack.Coll.Count.ToString)
                e.Message = "Excel processing segment " & i.ToString & " of " & ExcelPack.Coll.Count.ToString
                RaiseEvent NotifyForm(e)
            Next

            ' ___ Save and close the workbook
            oBook.SaveAs(OutputFullPath)
            oBook.Close()

            'oExcel.Quit()
            'ReleaseComObject(oCells)
            'ReleaseComObject(oSheet)
            'ReleaseComObject(oSheets)
            'ReleaseComObject(oBook)
            'ReleaseComObject(oBooks)
            'ReleaseComObject(oExcel)
            'oExcel = Nothing
            'oBooks = Nothing
            'oBook = Nothing
            'oSheets = Nothing
            'oSheet = Nothing
            'oCells = Nothing
            'System.GC.Collect()

            MyResults.Success = True
            MyResults.Value = "./TempData/" & OutputFileName
            Return MyResults

            '"./TempData/CCM_20090519_204900.xls"

        Catch ex As Exception
            MyResults.Success = False
            MyResults.Message = "Error #554.1: Excel ExportToExcel " & Replace(ex.Message, "'", "''")
            Return MyResults

        Finally
            Try
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
            Catch
            End Try
        End Try
    End Function

    Private Function LoadData(ByRef ExcelPackItem As ExcelPack_RptSupervisor.Item, ByRef RangeDataColl As Collection, ByRef oCells As Microsoft.Office.Interop.Excel.Range, ByRef oBook As Microsoft.Office.Interop.Excel.Workbook, ByVal ShowColumnNames As Boolean) As Results
        Dim i As Integer
        Dim dr As DataRow
        Dim ItemArray() As Object
        Dim dtRowNum As Integer
        Dim DateColNum As Integer
        Dim StartRow As Integer
        Dim Querypack As QueryPack
        Dim dt As DataTable
        Dim StartColumn As Integer
        Dim RangeDataDetailColl As Collection
        Dim ColumnName As String
        Dim StartColumnName As String
        Dim SumStatement As String
        Dim Range As String
        Dim MyResults As New Results

        Try

            ' ___ Unload the datatable from the ExcelSegment
            Querypack = cCommon.GetDTWithQuerypack(ExcelPackItem.Sql)
            If Not Querypack.Success Then
                MyResults.Success = False
                MyResults.Message = "Error #555a Excel.LoadData " & Querypack.TechErrMsg
                Return MyResults
            End If

            dt = Querypack.dt

            ' ___ Get the start column and row
            StartColumn = MyBase.ExcelColumnToNumber(RangeDataColl("xlColumnName")) + ExcelPackItem.SegmentOffset
            StartColumnName = MyBase.GetNumberToLetter(StartColumn, BasisEnum.One)
            StartRow = RangeDataColl("xlRowNum")
            If ShowColumnNames Then
                StartRow += 1
            End If

            ' ___ Output Column Headers
            If ShowColumnNames Then
                For i = 0 To dt.Columns.Count - 1
                    oCells(7, StartColumn + i) = dt.Columns(i).ColumnName
                Next
            End If

            ' ___ Output Data
            For dtRowNum = 0 To dt.Rows.Count - 1
                dr = dt.Rows.Item(dtRowNum)
                ItemArray = dr.ItemArray
                For i = 0 To ItemArray.GetUpperBound(0)
                    oCells(StartRow + dtRowNum, StartColumn + i) = ItemArray(i).ToString
                Next
            Next

            If dt.Rows.Count > 0 Then

                ' ___ Sum calculation
                If Not ExcelPackItem.SuppressTotalLabel Then
                    ColumnName = StartColumnName
                    For i = 0 To dt.Columns.Count - 1
                        SumStatement = "=SUM(" & ColumnName & StartRow.ToString & ":" & ColumnName & (StartRow + dt.Rows.Count - 1).ToString & ")"
                        oCells(7, StartColumn + i) = SumStatement
                        ColumnName = AddColumn(ColumnName)
                    Next
                End If

                ' ___ Right-justify summaries
                ColumnName = StartColumnName
                Range = ColumnName & "7:" & AddColumnThisNumber(ColumnName, dt.Columns.Count - 1) & "7"
                oCells.Range(Range).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlRight

                ' ___ Total label in name column
                If ExcelPackItem.SegmentType = SegmentType.Client AndAlso Not ExcelPackItem.SuppressTotalLabel Then
                    Range = StartColumnName & "7:" & StartColumnName & "7"
                    oCells.Range(Range).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                    oSheet.Range(Range).Font.Bold = True
                    oCells(7, StartColumn) = "Total"
                End If
            End If

            ' ___ Date and location
            If ExcelPackItem.Location = Nothing And ExcelPackItem.ReportDate = Nothing Then
            Else
                If ExcelPackItem.SegmentType = SegmentType.Client Then
                    RangeDataDetailColl = GetRangeData("Date", ExcelPackItem, oBook)
                    DateColNum = MyBase.ExcelColumnToNumber(RangeDataDetailColl("xlColumnName"))
                    oCells(1, DateColNum) = CType(ExcelPackItem.ReportDate, Date).ToString("M/dd/yyyy")
                    ' RangeDataDetailColl = GetRangeData("Location", ExcelPackItem, oBook)
                    'oCells(2, DateColNum) = ExcelPackItem.Location
                End If
            End If

            MyResults.Success = True
            Return MyResults

        Catch ex As Exception
            MyResults.Success = False
            MyResults.Message = "Error #555b Excel.LoadData " & Replace(ex.Message, "'", "''")
            Return MyResults
            'Throw New Exception("Error #555: ExcelOut DumpData " & ex.Message, ex)
        End Try
    End Function

    Private Function AddColumnThisNumber(ByVal ColName As String, ByVal Num As Integer) As String
        Dim Count As Integer
        Dim ThisColName As String

        Try

            If Num = 0 Then
                Return ColName
            End If

            ThisColName = ColName
            Do
                Count += 1
                ThisColName = AddColumn(ThisColName)
                If Count = Num Then
                    Return ThisColName
                End If

                If Count > 500 Then
                    Throw New Exception("Error #558a: ExcelOut AddColumnThisNum. Count exceeds 500")
                End If
            Loop

        Catch ex As Exception
            Throw New Exception("Error #558b: ExcelOut AddColumnThisNumber " & ex.Message, ex)
        End Try
    End Function

    Private Function AddColumn(ByVal ColName As String) As String
        Dim i As Integer
        Dim MaxLength As Integer = 5
        Dim Rec_(MaxLength - 1) As String
        Dim Results As String
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

    Private Function GetRangeData(ByVal RangeType As String, ByRef ExcelPackItem As ExcelPack_RptSupervisor.Item, ByRef oBook As Microsoft.Office.Interop.Excel.Workbook) As Collection
        Dim i As Integer
        Dim xlRangeName As String
        Dim SearchRangeName As String
        Dim Address As Object
        Dim Box1() As String
        Dim Box2() As String
        Dim SheetName As String
        Dim SheetNum As Integer
        Dim xlColumnName As String
        Dim xlRowNum As String
        Dim Coll As New Collection

        Try


            If ExcelPackItem.ClientID = "C3" Then
                SearchRangeName = "CCC_" & RangeType
            Else
                SearchRangeName = Replace(ExcelPackItem.ClientID, "|", "") & "_" & RangeType
            End If

            SearchRangeName = SearchRangeName.ToUpper

            For i = 1 To oBook.Names.Count
                xlRangeName = oBook.Names.Item(i).Name.ToUpper
                If SearchRangeName = xlRangeName Then
                    Address = oBook.Names.Item(i)._Default
                    Box1 = Split(Address, "!")
                    SheetName = Box1(0).Substring(1)
                    SheetName = Replace(SheetName, "'", "")
                    Box2 = Split(SheetName, "Sheet")
                    SheetNum = Box2(1)
                    Box1 = Split(Box1(1), "$")
                    xlColumnName = Box1(1)
                    xlRowNum = Box1(2)
                    Exit For
                End If
            Next

            If xlColumnName = Nothing Then
                Throw New Exception("Error #100.1: ExcelOut GetRangeData Cannot find range named " & SearchRangeName)
            End If



            If RangeType = "Anchor" Then
                xlRowNum += 2
            End If

            Coll.Add(SheetNum, "SheetNum")
            Coll.Add(xlRowNum, "xlRowNum")
            Coll.Add(xlColumnName, "xlColumnName")
            Return Coll

        Catch ex As Exception
            'Throw New Exception("Error #557: ExcelOut GetRangeData " & ex.Message, ex)
            cReport.Report("Excel.GetRangeData  #100 " & Replace(ex.Message, "'", "''"), Report.ReportTypeEnum.LogError)
        End Try
    End Function

    'Private Sub SecondDumpData(ByRef ExcelSegment As ExcelPack.Segment, ByVal oCells As Microsoft.Office.Interop.Excel.Range, ByVal ShowColumnNames As Boolean)
    '    Dim dr As DataRow
    '    Dim ItemArray() As Object
    '    Dim RowNum As Integer
    '    Dim ColNum As Integer
    '    Dim FirstDataRow As Integer
    '    Dim dt As DataTable
    '    Dim StartColumn As Integer
    '    Dim DateLocationColumn As String

    '    Try

    '        dt = ExcelSegment.dt
    '        StartColumn = ExcelSegment.StartColumn

    '        If ShowColumnNames Then
    '            FirstDataRow = 8
    '        Else
    '            FirstDataRow = 7
    '        End If


    '        ' ___ Output Column Headers
    '        If ShowColumnNames Then
    '            For ColNum = 0 To dt.Columns.Count - 1
    '                oCells(7, StartColumn + ColNum + 1) = dt.Columns(ColNum).ColumnName
    '            Next
    '        End If

    '        ' ___ Output Data
    '        For RowNum = 0 To dt.Rows.Count - 1
    '            dr = dt.Rows.Item(RowNum)
    '            ItemArray = dr.ItemArray
    '            For ColNum = 0 To ItemArray.GetUpperBound(0)
    '                oCells(FirstDataRow + RowNum, StartColumn + ColNum + 1) = ItemArray(ColNum).ToString
    '            Next
    '        Next

    '        ' ___ Date and location
    '        If ExcelSegment.SegmentType = SegmentType.Client Then
    '            ColNum = ColumnToNumber(ExcelSegment.DateLocationColumn)
    '            oCells(1, ColNum) = CType(ExcelSegment.ReportDate, Date).ToString("M/dd/yyyy")
    '            oCells(2, ColNum) = cEnviro.LoginLocationID
    '        End If

    '    Catch ex As Exception
    '        Throw New Exception("Error #555: ExcelOut DumpData " & ex.Message)
    '    End Try

    'End Sub

    'Private Function ColumnToNumber(ByVal ColumnLtr As String) As Integer
    '    Dim LeftLetter As String
    '    Dim RightLetter As String
    '    Dim LeftWorking As Integer
    '    Dim RightWorking As Integer

    '    Try
    '        ColumnLtr = ColumnLtr.ToUpper
    '        If ColumnLtr.Length = 1 Then
    '            RightLetter = ColumnLtr
    '        Else
    '            LeftLetter = ColumnLtr.Substring(0, 1)
    '            RightLetter = ColumnLtr.Substring(1, 1)
    '        End If
    '        RightWorking = Asc(RightLetter) - 64
    '        If LeftLetter <> Nothing Then
    '            LeftWorking = 26 * (Asc(LeftLetter) - 64)
    '        End If
    '        Return LeftWorking + RightWorking

    '    Catch ex As Exception
    '        ' Throw New Exception("Error #556: ExcelOut ColumnToNumber " & ex.Message, ex)
    '        cReport.Report("Excel.ColumnToNumber  #100 " & ex.Message, Report.ReportTypeEnum.LogError)
    '    End Try
    'End Function

    'Private Sub VeryFirstDumpData(ByVal dt As DataTable, ByVal oCells As Microsoft.Office.Interop.Excel.Range, ByVal ShowColumnNames As Boolean, ByVal StartColumn As Integer)
    '    Dim dr As DataRow
    '    Dim ItemArray() As Object
    '    Dim RowNum As Integer
    '    Dim ColNum As Integer
    '    Dim FirstDataRow As Integer

    '    Try

    '        If ShowColumnNames Then
    '            FirstDataRow = 8
    '        Else
    '            FirstDataRow = 7
    '        End If


    '        ' ___ Output Column Headers
    '        If ShowColumnNames Then
    '            For ColNum = 0 To dt.Columns.Count - 1
    '                oCells(7, StartColumn + ColNum + 1) = dt.Columns(ColNum).ColumnName
    '            Next
    '        End If

    '        ' ___ Output Data
    '        For RowNum = 0 To dt.Rows.Count - 1
    '            dr = dt.Rows.Item(RowNum)
    '            ItemArray = dr.ItemArray
    '            For ColNum = 0 To ItemArray.GetUpperBound(0)
    '                oCells(FirstDataRow + RowNum, StartColumn + ColNum + 1) = ItemArray(ColNum).ToString
    '            Next
    '        Next


    '        '' ___ Output Column Headers
    '        'If ShowColumnNames Then
    '        '    For ColNum = 0 To dt.Columns.Count - 1
    '        '        oCells(7, ColNum + 1) = dt.Columns(ColNum).ColumnName
    '        '    Next
    '        'End If

    '        '' ___ Output Data
    '        'For RowNum = 0 To dt.Rows.Count - 1
    '        '    dr = dt.Rows.Item(RowNum)
    '        '    ItemArray = dr.ItemArray
    '        '    For ColNum = 0 To ItemArray.GetUpperBound(0)
    '        '        oCells(FirstDataRow + RowNum, ColNum + 1) = ItemArray(ColNum).ToString
    '        '    Next
    '        'Next

    '    Catch ex As Exception
    '        Throw New Exception("Error #555: ExcelOut DumpData " & ex.Message)
    '    End Try

    'End Sub
End Class
