# CompuMaster.Data.XlsEpplus
Read datatables from/write to Microsoft Excel files (XLSX) with **1 line of code**

## Sample for read data from XLSX files to a System.Data.DataTable
```vb.net
  Dim t as System.Data.DataTable = CreateTestData()
  t = CompuMaster.Data.XlsEpplus.ReadDataSetFromXlsFile(TempFile, True)
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
