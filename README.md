# CompuMaster.Data.XlsEpplus
Read datatables from/write to Microsoft Excel files (XLSX) with **1 line of code**

[![Github Release](https://img.shields.io/github/release/CompuMasterGmbH/CompuMaster.Data.XlsEpplus.svg?maxAge=2592000&label=GitHub%20Release)](https://github.com/CompuMasterGmbH/CompuMaster.Data.XlsEpplus/releases) 
[![NuGet CompuMaster.Data.XlsEpplus](https://img.shields.io/nuget/v/CompuMaster.Data.XlsEpplus.svg?label=NuGet%20CM.Data.XlsEpplus)](https://www.nuget.org/packages/CompuMaster.Data.XlsEpplus/)

## Sample for read data from XLSX files to a System.Data.DataTable
```vb.net
  Dim t As System.Data.DataTable = CompuMaster.Data.XlsEpplus.ReadDataTableFromXlsFile(TempFile)
  Dim ds As System.Data.DataSet = CompuMaster.Data.XlsEpplus.ReadDataSetFromXlsFile(TempFile, True)
  Dim SheetNames As String() = CompuMaster.Data.XlsEpplus.ReadSheetNamesFromXlsFile(TempFile)
  Dim t2 As System.Data.DataTable = CompuMaster.Data.XlsEpplus.ReadDataTableFromXlsFile(TempFile, SheetNames(2))
```

## Sample for writing data to XLSX files
```vb.net
Sub Test()
  Dim t as System.Data.DataTable = CreateTestData()
  'Write test data
  CompuMaster.Data.XlsEpplus.WriteDataTableToXlsFile(Nothing, TempFile, data, "test")
End Sub

Function CreateTestData() As System.Data.DataTable
  Dim data As New DataTable("testtable")
  data.Columns.Add(New DataColumn("some values", GetType(String)))
  data.Columns.Add(New DataColumn("a second column", GetType(String)))
  Dim row As DataRow
  row = data.NewRow
  row("some values") = "this is a Test!"
  row("a second column") = "2 columns and 2 rows"
  data.Rows.Add(row)
  row = data.NewRow
  row("some values") = "this is the last line"
  row("a second column") = "this is the last cell" & System.Environment.NewLine & "2nd line" & System.Environment.NewLine & "3rd line" & System.Environment.NewLine & "4th line" & vbTab & "with a tab char"
  data.Rows.Add(row)
  Return data
End Function
```
