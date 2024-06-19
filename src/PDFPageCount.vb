Imports System
Imports System.IO
Imports System.Text
Imports iTextSharp.text.pdf

Public Class PDFPageCount

    Private _A3 As Integer = 0
    Private _A4 As Integer = 0
    Private _Total As Integer = 0

    Public ReadOnly Property Total As Integer
        Get
            Return _Total
        End Get
    End Property
    Public ReadOnly Property A3 As Integer
        Get
            Return _A3
        End Get
    End Property
    Public ReadOnly Property A4 As Integer
        Get
            Return _A4
        End Get
    End Property
    Public Sub New(inFile As String)
        GetSize(inFile)
    End Sub
    Private Sub GetSize(inFile As String)
        Dim info As PdfReader = New PdfReader(inFile)
        _Total = info.NumberOfPages
        For i As Integer = 1 To (_Total + 1) - 1
            If info.GetPageSize(i).Right > 1000.0! OrElse info.GetPageSize(i).Top > 1000.0! Then
                _A3 += 1
            End If
        Next i
        _A4 = _Total - _A3
    End Sub
End Class
