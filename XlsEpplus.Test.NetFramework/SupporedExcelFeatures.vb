Imports NUnit.Framework

Namespace CompuMaster.Test.Data

    <TestFixture()> Public Class SupportedExcelFeatures

        <Test> Sub TableWithCalculatedCellsTargettingTableColumnsByColumnName()
            Dim file As String = GlobalTestSetup.PathToTestFiles("testfiles\tablevalues.xlsx")
            'Dim ds As DataSet = CompuMaster.Data.XlsEpplus.ReadDataSetFromXlsFile(file, True)
            Dim dt As DataTable
            dt = CompuMaster.Data.XlsEpplus.ReadDataTableFromXlsFile(file)
            Dim ColCount As Integer = dt.Columns.Count
            dt = CompuMaster.Data.DataTables.CopyDataTableWithSubsetOfRows(dt, 5)
            Dim FormattedTextTable As String = CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt)
            Console.WriteLine(FormattedTextTable)
            If FormattedTextTable.Contains("NotImplementedException") Then
                Assert.Fail("Found cell content with ""NotImplementedException""")
            End If
        End Sub

    End Class

End Namespace