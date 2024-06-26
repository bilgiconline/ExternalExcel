Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports System.IO

Namespace ExternalExcel

    Public Class ExternalExcelClass
        <InterfaceType(ComInterfaceType.InterfaceIsIDispatch)> Public Interface IExternalExcel
            Property ApplicationVisible As Boolean
            Property DisplayErrors As Boolean
            Sub CreateNewWorkbook()
            Sub CloseWorkbook()
            Function OpenWorkbook(FilePath As String) As String
            Function GetWorksheet(SheetsName As String) As String
            Function GetValue(ColumnName As String, Row As Integer) As String
            Function UsedRangeCount() As Long
            Sub Dispose()

        End Interface
        <ClassInterface(ClassInterfaceType.None)> Public Class ExternalExcel
            Implements IExternalExcel
            Declare Function SetProcessWorkingSetSize Lib "kernel32.dll" (ByVal process As IntPtr, ByVal minimumWorkingSetSize As Integer, ByVal maximumWorkingSetSize As Integer) As Integer
            Dim o_Excel As New Excel.Application
            Dim Wb As Workbook
            Dim Ws As Worksheet
            Private disposedValue As Boolean
            Property ApplicationVisible As Boolean Implements IExternalExcel.ApplicationVisible
                Get
                    ApplicationVisible = o_Excel.Application.Visible
                End Get
                Set(value As Boolean)
                    o_Excel.Application.Visible = value
                End Set
            End Property

            Property DisplayErrors As Boolean Implements IExternalExcel.DisplayErrors
                Get
                    DisplayErrors = o_Excel.Application.DisplayAlerts
                End Get
                Set(value As Boolean)
                    o_Excel.Application.DisplayAlerts = value
                End Set
            End Property

            Sub CreateNewWorkbook() Implements IExternalExcel.CreateNewWorkbook
                Wb = o_Excel.Workbooks.Add()
            End Sub

            Sub CloseWorkbook() Implements IExternalExcel.CloseWorkbook
                Try
                    Wb.Close()
                    o_Excel.Quit()
                Catch ex As Exception

                End Try
            End Sub
            Function OpenWorkbook(FilePath As String) As String Implements IExternalExcel.OpenWorkbook
                If File.Exists(FilePath) = True Then
                    Try
                        If o_Excel Is Nothing Then o_Excel = New Application
                        Wb = o_Excel.Workbooks.Open(FilePath)
                        Return "File opened."
                    Catch ex As Exception
                        Return "File error."
                    End Try
                Else
                    Return "File was not found. Please check file path."
                End If
            End Function
            Function GetWorksheet(SheetsName As String) As String Implements IExternalExcel.GetWorksheet
                If Wb Is Nothing Or Wb.Name = "" Then
                    Return "Not found any workbook. Please open a work file."
                Else
                    For k As Integer = 0 To Wb.Sheets.Count - 1
                        If Wb.Sheets(k).Name = SheetsName Then
                            Ws = Wb.Sheets(k)
                            Return "Selected worksheet."
                            Exit For
                        End If
                    Next
                    Try
                        If Ws Is Nothing Or Ws.Name = "" Then
                            Return "Error, cant selected worksheet."
                        End If
                    Catch ex As Exception
                        Return "Error, cant selected worksheet."
                    End Try
                End If
            End Function
            Function GetValue(ColumnName As String, Row As Integer) As String Implements IExternalExcel.GetValue
                If Ws.Name Is Nothing Or Ws.Name = "" Then
                    Return "Error, cant selected worksheet."
                Else
                    If ColumnName = "" Then
                        Return "Column was empty."
                        Exit Function
                    End If
                    If Row = "" Then
                        Return "Row was empty."
                        Exit Function
                    End If
                    Return CStr(Ws.Range(ColumnName & Row).Value)
                End If
            End Function
            Function UsedRangeCount() As Long Implements IExternalExcel.UsedRangeCount
                If Ws.Name Is Nothing Or Ws.Name = "" Then
                    Return "Error, cant selected worksheet."
                Else
                    Return Ws.UsedRange.Rows.Count
                End If
            End Function
            Protected Overridable Sub Dispose(disposing As Boolean)
                If Not Me.disposedValue Then
                    If disposing Then
                        Try
                            o_Excel.Quit()
                        Catch ex As Exception

                        End Try
                    End If
                End If
                Me.disposedValue = True
            End Sub

            Sub Dispose() Implements IExternalExcel.Dispose
                Dispose(True)
                GC.SuppressFinalize(Me)
            End Sub
        End Class
    End Class
End Namespace