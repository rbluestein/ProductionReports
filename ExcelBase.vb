Public Class ExcelBase
    Protected Function ExcelColumnToNumber(ByVal ColumnLtr As String) As Integer
        Dim LeftLetter As String
        Dim RightLetter As String
        Dim LeftWorking As Integer
        Dim RightWorking As Integer
        Dim Report As New Report

        Try
            ColumnLtr = ColumnLtr.ToUpper
            If ColumnLtr.Length = 1 Then
                RightLetter = ColumnLtr
            Else
                LeftLetter = ColumnLtr.Substring(0, 1)
                RightLetter = ColumnLtr.Substring(1, 1)
            End If
            RightWorking = Asc(RightLetter) - 64
            If LeftLetter <> Nothing Then
                LeftWorking = 26 * (Asc(LeftLetter) - 64)
            End If
            Return LeftWorking + RightWorking

        Catch ex As Exception
            ' Throw New Exception("Error #556: ExcelOut ColumnToNumber " & ex.Message, ex)
            Report.Report("ExcelBase.ExcelColumnToNumber  #100 " & ex.Message, Report.ReportTypeEnum.LogError)
        End Try
    End Function

    Protected Function GetNumberToLetter(ByVal Number As Integer, ByVal Basis As BasisEnum) As String
        Dim Results As String

        Try

            ' ___ Convert to 0-basis to make modulo/divide cleaner. 
            If Basis = BasisEnum.One Then
                Number = Number - 1
            End If

            ' ___ Select return value based on invalid/one-char/two-char input.
            If Number < 0 Or Number >= 27 * 26 Then
                ' ___ Return special sentinel value if out of range.
                Results = "Invalid input of " & Number.ToString
            Else
                ' ___Single char, just get the letter.
                If Number < 26 Then
                    Results = Chr(Number + 65)
                Else
                    ' ___ Double char, get letters based on integer divide and modulus.
                    Results = Chr(Number \ 26 + 64) + Chr(Number Mod 26 + 65)
                End If
            End If

            Return Results
        Catch ex As Exception
            Throw New Exception("Error #2828: ExcelHandler GetNumberToLetter " & ex.Message, ex)
        End Try
    End Function

    Public Class ExcelAddress
        Private cRangeName As String
        Private cRowNum As Integer
        Private cColumnLtr As String
        Private cRowOffset As Integer
        Private cColumnOffset As Integer
        Private cSheetNum As Integer

        Public Property RangeName() As String
            Get
                Return cRangeName
            End Get
            Set(ByVal Value As String)
                cRangeName = Value
            End Set
        End Property
        Public Property RowNum() As Integer
            Get
                Return cRowNum
            End Get
            Set(ByVal Value As Integer)
                cRowNum = Value
            End Set
        End Property
        Public Property ColumnLtr() As String
            Get
                Return cColumnLtr
            End Get
            Set(ByVal Value As String)
                cColumnLtr = Value
            End Set
        End Property
        Public Property RowOffset() As Integer
            Get
                Return cRowOffset
            End Get
            Set(ByVal Value As Integer)
                cRowOffset = Value
            End Set
        End Property
        Public Property ColumnOffset() As Integer
            Get
                Return cColumnOffset
            End Get
            Set(ByVal Value As Integer)
                cColumnOffset = Value
            End Set
        End Property
        Public Property SheetNum() As Integer
            Get
                Return cSheetNum
            End Get
            Set(ByVal Value As Integer)
                cSheetNum = Value
            End Set
        End Property
    End Class
End Class
