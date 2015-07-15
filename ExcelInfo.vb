Imports System.Runtime.InteropServices.Marshal

Public Class ExcelInfo
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

    Public Function SetFill(ByVal SheetNum As Integer, ByVal RowNum As Integer) As Microsoft.Office.Interop.Excel.Application
        oSheet = CType(oSheets.Item(SheetNum), Microsoft.Office.Interop.Excel.Worksheet)
        With oSheet.Range(RowNum & ":" & RowNum).Interior
            .ColorIndex = 6
        End With
        Return oExcel
    End Function


    Public Sub New(ByVal WorkbookName As String)
        Try
            cEnviro = gEnviro
            oExcel = New Microsoft.Office.Interop.Excel.Application

            ' ___ Identify the template and the output files 
            cAppPath = cEnviro.GetAppPath
            cTemplateFullPath = cAppPath & "\Templates\" & WorkbookName
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
            cReport.Report("Excel_Generic.New  #1202 " & ex.Message, Report.ReportTypeEnum.LogError)
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
        Catch
        End Try
    End Sub

    Public Function ExportFieldToExcel(ByVal Value As String, ByRef ExcelAddress As ExcelAddress, Optional ByRef ActiveExcel As Microsoft.Office.Interop.Excel.Application = Nothing) As Results
        Dim Results As Results
        Results = ExportToExcel(ExportSourceEnum.Field, Value, ExcelAddress, ActiveExcel)
        Return Results
    End Function

    Public Function ExportCollectionToExcel(ByVal Coll As Collection, ByRef ExcelAddress As ExcelAddress, Optional ByRef ActiveExcel As Microsoft.Office.Interop.Excel.Application = Nothing) As Results
        Dim Results As Results
        Results = ExportToExcel(ExportSourceEnum.Collection, Coll, ExcelAddress, ActiveExcel)
        Return Results
    End Function

    Public Function ExportTableToExcel(ByVal dt As DataTable, ByRef ExcelAddress As ExcelAddress, Optional ByRef ActiveExcel As Microsoft.Office.Interop.Excel.Application = Nothing) As Results
        Dim Results As Results
        Results = ExportToExcel(ExportSourceEnum.Table, dt, ExcelAddress, ActiveExcel)
        Return Results
    End Function

    Public Function ExportToExcel(ByVal ExportSource As ExportSourceEnum, ByVal Value As Object, ByRef ExcelAddress As ExcelAddress, ByRef ActiveExcel As Microsoft.Office.Interop.Excel.Application) As Results
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
                '  RangeDataColl = GetRangeData(ExcelAddress.RangeName, oBook)
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
                    oCells(StartRow, StartColumn) = Value

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
            MyResults.Message = "Error #594: Excel ExportToExcel " & ex.Message
            Return MyResults
        Finally
            If ErrorInd Then
                Try
                    Finish(oExcel)
                Catch
                End Try
            End If
        End Try
    End Function


    Public Function GetRangeData() As Collection
        Dim Coll As Collection
        Coll = GetRangeData(oBook)
        Return Coll
    End Function

    Private Function GetRangeData(ByRef oBook As Microsoft.Office.Interop.Excel.Workbook) As Collection
        Dim i As Integer
        Dim xlRangeName As String
        Dim Address As Object
        Dim Box1() As String
        Dim SheetName As String
        Dim CollColl As New Collection
        Dim Coll As Collection
        Dim StartRow As Integer
        Dim StartColumn As String
        Dim Text As String

        Try

            For i = 1 To oBook.Names.Count

                Try

                    Coll = New Collection

                    xlRangeName = oBook.Names.Item(i).Name
                    Address = oBook.Names.Item(i)._Default
                    Box1 = Split(Address, "!")
                    SheetName = Box1(0).Substring(1)
                    SheetName = Replace(SheetName, "'", "")
                    Box1 = Split(Box1(1), "$")

                    StartRow = Replace(Box1(2), ":", "")
                    StartColumn = Box1(1)

                    Coll.Add(xlRangeName, "RangeName")
                    Coll.Add(SheetName, "SheetName")
                    Coll.Add(Replace(Box1(2), ":", ""), "StartRow")
                    Coll.Add(Box1(1), "StartCol")

                    If Box1.GetUpperBound(0) = 4 Then
                        Coll.Add(Box1(4), "EndRow")
                        Coll.Add(Box1(3), "EndCol")
                    End If

                    Dim Sheet As Microsoft.Office.Interop.Excel.Worksheet
                    Dim Cells As Microsoft.Office.Interop.Excel.Range

                    For Each Sheet In oBook.Worksheets
                        If Sheet.Name.ToLower = SheetName.ToLower Then
                            Coll.Add(Sheet.Index, "SheetNum")
                            Cells = Sheet.Cells
                            If IsDBNull(Cells(StartRow, MyBase.ExcelColumnToNumber(StartColumn)).Text) Then
                                Text = String.Empty
                            Else
                                Text = Cells(StartRow, MyBase.ExcelColumnToNumber(StartColumn)).Text
                            End If
                            Coll.Add(Text, "Text")
                            Exit For
                        End If
                    Next

                    CollColl.Add(Coll, xlRangeName)

                Catch ex As Exception
                End Try

            Next

            For i = 1 To CollColl.Count
                Try
                    System.Diagnostics.Debug.WriteLine(CollColl(i)("RangeName") & vbTab & CollColl(i)("StartCol") & CollColl(i)("StartRow") & vbTab & CollColl(i)("Text"))
                Catch
                End Try
            Next

            Return CollColl

        Catch ex As Exception
            ' Throw New Exception("Error #557: ExcelOut GetRangeData " & ex.Message, ex)
            'cReport.Report("Excel_Generic.GetRangeData  #100b RangeName: " & RangeName & "   " & ex.Message, Report.ReportTypeEnum.LogError)
        End Try
    End Function
End Class

