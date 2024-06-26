Imports externalExcel
Module test
    Sub New()
        Dim ExcelApp As New ExternalExcel.ExternalExcelClass.ExternalExcel
        ExcelApp.DisplayErrors = False
        ExcelApp.ApplicationVisible = False
        ExcelApp.OpenWorkbook(".\your_file_path\example.xlsx")
        ExcelApp.GetWorksheet("Sheet1")
        Debug.Print(ExcelApp.GetValue("A", 5))
        Debug.Print(ExcelApp.UsedRangeCount)
        ExcelApp.CloseWorkbook()
        ExcelApp.Dispose()
        ExcelApp = Nothing
    End Sub
End Module
