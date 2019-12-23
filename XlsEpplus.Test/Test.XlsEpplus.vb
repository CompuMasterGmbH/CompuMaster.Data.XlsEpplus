Imports NUnit.Framework

Namespace CompuMaster.Test.Data

    <TestFixture()> Public Class XlsEpplus

        Private ReadOnly Property TempFile() As String
            Get
                Static _TempFile As String
                If _TempFile Is Nothing Then
                    _TempFile = System.IO.Path.GetTempFileName & ".xlsx"
                    Console.WriteLine(_TempFile)
                End If
                Return _TempFile
            End Get
        End Property

        <Test()> Public Sub SaveAndReadSimple()
            'Prepare test data
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
            'Write test data
            CompuMaster.Data.XlsEpplus.WriteDataTableToXlsFile(Nothing, TempFile, data, "test")

            'Read and compare written test data
            '==================================

            'the number of sheets should be one (because we've created a new XLS file)
            Assert.AreEqual(1, CompuMaster.Data.XlsEpplus.ReadDataSetFromXlsFile(TempFile, False).Tables.Count, "SaveAndReadSimple #01")

            'read the existing file, auto-detect column-types, take datatable and compare it with the written data: it should be always the same (or must be argumented and discussed with Jochen why it isn't)
            'the number of columns and rows should be always 2
            Dim ReReadData As DataTable
            ReReadData = CompuMaster.Data.XlsEpplus.ReadDataTableFromXlsFile(TempFile, "test", True)
            Assert.AreEqual("test", ReReadData.TableName, "SaveAndReadSimple #05")
            Assert.AreEqual(2, ReReadData.Columns.Count, "SaveAndReadSimple #10")
            Assert.AreEqual("some values", ReReadData.Columns(0).ColumnName, "SaveAndReadSimple #11")
            Assert.AreEqual("a second column", ReReadData.Columns(1).ColumnName, "SaveAndReadSimple #12")
            Assert.AreEqual(2, ReReadData.Rows.Count, "SaveAndReadSimple #20")
            Assert.AreEqual("this is a Test!", ReReadData.Rows(0)(0), "SaveAndReadSimple #21")
            Assert.AreEqual("2 columns and 2 rows", ReReadData.Rows(0)(1), "SaveAndReadSimple #22")
            Assert.AreEqual("this is the last line", ReReadData.Rows(1)(0), "SaveAndReadSimple #23")
            Assert.AreEqual("this is the last cell" & System.Environment.NewLine & "2nd line" & System.Environment.NewLine & "3rd line" & System.Environment.NewLine & "4th line" & vbTab & "with a tab char", ReReadData.Rows(1)(1), "SaveAndReadSimple #24")
        End Sub

        <Test()> Public Sub SaveAndReadUnicode()
            'Prepare test data
            Dim data As New DataTable("testtable")
            data.Columns.Add(New DataColumn("some values", GetType(String)))
            Dim row As DataRow
            row = data.NewRow
            row("some values") = "ПК дома"
            data.Rows.Add(row)
            row = data.NewRow
            row("some values") = "^!§$%&/()=?´`~+*#'-_.:,;<>|\ÄÖÜäöü@€"
            data.Rows.Add(row)
            row = data.NewRow
            row("some values") = "セキュリティ更新プログラム"
            data.Rows.Add(row)
            row = data.NewRow
            row("some values") = "보안 비디오"
            data.Rows.Add(row)
            row = data.NewRow
            row("some values") = "Preuzimanje predložaka na Office Online"
            data.Rows.Add(row)
            row = data.NewRow
            row("some values") = "من عنده تأشيرة سكن في أيّ من دول مجلس التعاون الخليجي"
            data.Rows.Add(row)
            'Write test data
            CompuMaster.Data.XlsEpplus.WriteDataTableToXlsFile(Nothing, TempFile, data, "test")

            'Read and compare written test data
            '==================================

            'read the existing file, auto-detect column-types, take datatable and compare it with the written data: it should be always the same (or must be argumented and discussed with Jochen why it isn't)
            'the number of columns and rows should be always 2
            Dim ReReadData As DataTable
            ReReadData = CompuMaster.Data.XlsEpplus.ReadDataTableFromXlsFile(TempFile, "test", True)
            Assert.AreEqual("test", ReReadData.TableName, "SaveAndReadUnicode #05")
            Assert.AreEqual(1, ReReadData.Columns.Count, "SaveAndReadUnicode #10")
            Assert.AreEqual("some values", ReReadData.Columns(0).ColumnName, "SaveAndReadUnicode #11")
            Assert.AreEqual(6, ReReadData.Rows.Count, "SaveAndReadUnicode #20")
            Assert.AreEqual("ПК дома", ReReadData.Rows(0)(0), "SaveAndReadUnicode #21")
            Assert.AreEqual("^!§$%&/()=?´`~+*#'-_.:,;<>|\ÄÖÜäöü@€", ReReadData.Rows(1)(0), "SaveAndReadUnicode #22")
            Assert.AreEqual("セキュリティ更新プログラム", ReReadData.Rows(2)(0), "SaveAndReadUnicode #23")
            Assert.AreEqual("보안 비디오", ReReadData.Rows(3)(0), "SaveAndReadUnicode #24")
            Assert.AreEqual("Preuzimanje predložaka na Office Online", ReReadData.Rows(4)(0), "SaveAndReadUnicode #25")
            Assert.AreEqual("من عنده تأشيرة سكن في أيّ من دول مجلس التعاون الخليجي", ReReadData.Rows(5)(0), "SaveAndReadUnicode #26")
        End Sub

        <Test()> Public Sub SaveAndReadExtraLargeFields()
            Const HundredChars As String = "1234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890"
            'Prepare test data
            Dim data As New DataTable("testtable")
            data.Columns.Add(New DataColumn("string", GetType(String)))
            Dim row As DataRow = data.NewRow
            row("string") = HundredChars & HundredChars & HundredChars & HundredChars
            data.Rows.Add(row)
            'Write test data
            CompuMaster.Data.XlsEpplus.WriteDataTableToXlsFile(Nothing, TempFile, data, "test")

            'Read and compare written test data
            '==================================

            'read the existing file, auto-detect column-types, take datatable and compare it with the written data: it should be always the same (or must be argumented and discussed with Jochen why it isn't)
            'the number of columns and rows should be always 2
            Dim ReReadData As DataTable
            ReReadData = CompuMaster.Data.XlsEpplus.ReadDataTableFromXlsFile(TempFile, "test", True)
            Assert.AreEqual(HundredChars & HundredChars & HundredChars & HundredChars, ReReadData.Rows(0)(0), "SaveAndReadExtraLargeFields #11")
        End Sub

        <Test()> Public Sub SaveAndReadExtraFieldsWithLineBreaks()
            'Prepare test data
            Dim data As New DataTable("testtable")
            data.Columns.Add(New DataColumn("string", GetType(String)))
            Dim row As DataRow = data.NewRow
            row("string") = "line 1" & ControlChars.Cr & "line 2" & ControlChars.Lf & "line 3" & ControlChars.CrLf & "line 4"
            data.Rows.Add(row)
            'Write test data
            CompuMaster.Data.XlsEpplus.WriteDataTableToXlsFile(Nothing, TempFile, data, "test")

            'Read and compare written test data
            '==================================

            'read the existing file, auto-detect column-types, take datatable and compare it with the written data: it should be always the same (or must be argumented and discussed with Jochen why it isn't)
            'the number of columns and rows should be always 2
            Dim ReReadData As DataTable
            ReReadData = CompuMaster.Data.XlsEpplus.ReadDataTableFromXlsFile(TempFile, "test", True)
            Assert.AreEqual("line 1" & System.Environment.NewLine & "line 2" & System.Environment.NewLine & "line 3" & System.Environment.NewLine & "line 4", ReReadData.Rows(0)(0), "SaveAndReadExtraLargeFieldsWithLineBreaks #11")
        End Sub

        <Test()> Public Sub SaveAndReadDBNull()
            'Prepare test data
            Dim data As New DataTable("testtable")
            data.Columns.Add(New DataColumn("some values", GetType(String)))
            data.Columns.Add(New DataColumn("string", GetType(String)))
            data.Columns.Add(New DataColumn("int16", GetType(Int16)))
            data.Columns.Add(New DataColumn("int32", GetType(Int32)))
            data.Columns.Add(New DataColumn("int64", GetType(Int64)))
            data.Columns.Add(New DataColumn("boolean", GetType(Boolean)))
            data.Columns.Add(New DataColumn("object", GetType(Object)))
            data.Columns.Add(New DataColumn("datetime", GetType(DateTime)))
            data.Columns.Add(New DataColumn("double", GetType(Double)))
            Dim row As DataRow = data.NewRow
            row("some values") = "this is a DBNull-Test!"
            row("string") = DBNull.Value
            row("int16") = DBNull.Value
            row("int32") = DBNull.Value
            row("int64") = DBNull.Value
            row("boolean") = DBNull.Value
            row("object") = DBNull.Value
            row("datetime") = DBNull.Value
            row("double") = DBNull.Value
            data.Rows.Add(row)
            'Write test data
            CompuMaster.Data.XlsEpplus.WriteDataTableToXlsFile(Nothing, TempFile, data, "test")

            'Read and compare written test data
            '==================================

            'read the existing file, auto-detect column-types, take datatable and compare it with the written data: it should be always the same (or must be argumented and discussed with Jochen why it isn't)
            'the number of columns and rows should be always 2
            Dim ReReadData As DataTable
            ReReadData = CompuMaster.Data.XlsEpplus.ReadDataTableFromXlsFile(TempFile, "test", True)
            Assert.AreEqual("test", ReReadData.TableName, "SaveAndReadDBNull #05")
            Assert.AreEqual(9, ReReadData.Columns.Count, "SaveAndReadDBNull #10")
            Assert.AreEqual("some values", ReReadData.Columns(0).ColumnName, "SaveAndReadDBNull #11")
            Assert.AreEqual(1, ReReadData.Rows.Count, "SaveAndReadDBNull #20")
            Assert.AreEqual("this is a DBNull-Test!", ReReadData.Rows(0)(0), "SaveAndReadDBNull #21")
            Assert.AreEqual(DBNull.Value, ReReadData.Rows(0)(1), "SaveAndReadDBNull #22")
            Assert.AreEqual(DBNull.Value, ReReadData.Rows(0)(2), "SaveAndReadDBNull #23")
            Assert.AreEqual(DBNull.Value, ReReadData.Rows(0)(3), "SaveAndReadDBNull #24")
            Assert.AreEqual(DBNull.Value, ReReadData.Rows(0)(4), "SaveAndReadDBNull #25")
            Assert.AreEqual(DBNull.Value, ReReadData.Rows(0)(5), "SaveAndReadDBNull #26")
            Assert.AreEqual(DBNull.Value, ReReadData.Rows(0)(6), "SaveAndReadDBNull #27")
            Assert.AreEqual(DBNull.Value, ReReadData.Rows(0)(7), "SaveAndReadDBNull #28")
            Assert.AreEqual(DBNull.Value, ReReadData.Rows(0)(8), "SaveAndReadDBNull #29")
        End Sub

        <Test(), Ignore("ToBeImplemented after Epplus bug has been fixed, see https://github.com/JanKallman/EPPlus/issues/573")> Public Sub SaveAndReadDoubleSpecials()
            'Prepare test data
            Dim data As New DataTable("testtable")
            data.Columns.Add(New DataColumn("doubleNaN", GetType(Double)))
            data.Columns.Add(New DataColumn("doubleNegInf", GetType(Double)))
            data.Columns.Add(New DataColumn("doublePosInf", GetType(Double)))
            data.Columns.Add(New DataColumn("doubleEps", GetType(Double)))
            data.Columns.Add(New DataColumn("doubleVal", GetType(Double)))
            Dim row As DataRow
            row = data.NewRow
            row("doubleNaN") = Double.NaN
            row("doubleNegInf") = Double.NegativeInfinity
            row("doublePosInf") = Double.PositiveInfinity
            row("doubleEps") = Double.Epsilon
            row("doubleVal") = 54246723.14521
            data.Rows.Add(row)
            'Write test data
            CompuMaster.Data.XlsEpplus.WriteDataTableToXlsFile(Nothing, TempFile, data, "test")

            'Read and compare written test data
            '==================================

            'read the existing file, auto-detect column-types, take datatable and compare it with the written data: it should be always the same (or must be argumented and discussed with Jochen why it isn't)
            'the number of columns and rows should be always 2
            Dim ReReadData As DataTable
            ReReadData = CompuMaster.Data.XlsEpplus.ReadDataTableFromXlsFile(TempFile, "test", True)
            Console.WriteLine(ColumnDataTypesToPlainTextTableFixedColumnWidths(ReReadData))
            Assert.AreEqual("test", ReReadData.TableName, "SaveAndReadDoubleSpecials #05")
            Assert.AreEqual(5, ReReadData.Columns.Count, "SaveAndReadDoubleSpecials #10")
            For MyCounter As Integer = 0 To 4
                Assert.AreEqual(GetType(System.Double), ReReadData.Columns(MyCounter).DataType, "SaveAndReadDoubleSpecials #11 with col index " & MyCounter)
            Next
            Assert.AreEqual(GetType(Double), ReReadData.Columns(0).DataType, "SaveAndReadDoubleSpecials #12")
            Assert.AreEqual(1, ReReadData.Rows.Count, "SaveAndReadDoubleSpecials #20")
            Assert.AreEqual(Double.NaN, ReReadData.Rows(0)(0), "SaveAndReadDoubleSpecials #21")
            Assert.AreEqual(Double.PositiveInfinity, ReReadData.Rows(1)(0), "SaveAndReadDoubleSpecials #22") '#NUM! is considered as PositiveInfinity
            Assert.AreEqual(Double.PositiveInfinity, ReReadData.Rows(2)(0), "SaveAndReadDoubleSpecials #23") '#NUM! is considered as PositiveInfinity
            Assert.AreEqual(Double.PositiveInfinity, ReReadData.Rows(3)(0), "SaveAndReadDoubleSpecials #24") '#NUM! is considered as PositiveInfinity; roundings to just 0 in excel require a #NUM exception in excel
            Assert.AreEqual(54246723.14521, ReReadData.Rows(4)(0), "SaveAndReadDoubleSpecials #25")
        End Sub

        <Test()> Public Sub LastCellDetection()
            'Prepare test data
            Dim data As New DataTable("testtable")
            data.Columns.Add(New DataColumn("string", GetType(String)))
            Dim row As DataRow = data.NewRow
            row("string") = Nothing
            data.Rows.Add(row)
            row = data.NewRow
            row("string") = ""
            data.Rows.Add(row)
            row = data.NewRow
            row("string") = Nothing
            data.Rows.Add(row)
            row = data.NewRow
            row("string") = DBNull.Value
            data.Rows.Add(row)
            'Write test data
            CompuMaster.Data.XlsEpplus.WriteDataTableToXlsFile(Nothing, TempFile, data, "test")

            'Read and compare written test data
            '==================================

            'read the existing file, auto-detect column-types, take datatable and compare it with the written data: it should be always the same (or must be argumented and discussed with Jochen why it isn't)
            'the number of columns and rows should be always 2
            Dim ReReadData As DataTable
            ReReadData = CompuMaster.Data.XlsEpplus.ReadDataTableFromXlsFile(TempFile, "test")
            Assert.AreEqual(0, ReReadData.Rows.Count, "SaveAndReadEmptyStates #10") 'because last 4 lines only contains DBNull/nothing/empty string values
            Assert.AreEqual(1, ReReadData.Columns.Count, "SaveAndReadEmptyStates #11") 'but the column "string" has been defined by the column header
        End Sub

        <Test()> Public Sub SaveAndReadEmptyStates()
            'Prepare test data
            Dim data As New DataTable("testtable")
            data.Columns.Add(New DataColumn("string", GetType(String)))
            data.Columns.Add(New DataColumn("dummy", GetType(String)))
            Dim row As DataRow = data.NewRow
            row("string") = Nothing
            data.Rows.Add(row)
            row = data.NewRow
            row("string") = ""
            data.Rows.Add(row)
            row = data.NewRow
            row("string") = Nothing
            data.Rows.Add(row)
            row = data.NewRow
            row("string") = DBNull.Value
            row("dummy") = "lastCell" 'required to ensure the excel file is read completely to the end (otherwise, empty rows/columns would be truncated)
            data.Rows.Add(row)
            'Write test data
            CompuMaster.Data.XlsEpplus.WriteDataTableToXlsFile(Nothing, TempFile, data, "test")

            'Read and compare written test data
            '==================================

            'read the existing file, auto-detect column-types, take datatable and compare it with the written data: it should be always the same (or must be argumented and discussed with Jochen why it isn't)
            'the number of columns and rows should be always 2
            Dim ReReadData As DataTable
            ReReadData = CompuMaster.Data.XlsEpplus.ReadDataTableFromXlsFile(TempFile, "test")
            Assert.AreEqual(4, ReReadData.Rows.Count, "SaveAndReadEmptyStates #10") 'because last 2 lines only contains DBNull
            Assert.AreEqual(DBNull.Value, ReReadData.Rows(0)(0), "SaveAndReadEmptyStates #11")
            Assert.AreEqual("", ReReadData.Rows(1)(0), "SaveAndReadEmptyStates #12")
            Assert.AreEqual(DBNull.Value, ReReadData.Rows(2)(0), "SaveAndReadEmptyStates #13")
            Assert.AreEqual(DBNull.Value, ReReadData.Rows(3)(0), "SaveAndReadEmptyStates #14")
        End Sub

        <Test()> Public Sub SaveAndReadDataTypes()
            'Prepare test data
            Dim data As New DataTable("testtable")
            data.Columns.Add(New DataColumn("some values", GetType(String)))
            data.Columns.Add(New DataColumn("string", GetType(String)))
            data.Columns.Add(New DataColumn("int16", GetType(Int16)))
            data.Columns.Add(New DataColumn("int32", GetType(Int32)))
            data.Columns.Add(New DataColumn("int64", GetType(Int64)))
            data.Columns.Add(New DataColumn("boolean", GetType(Boolean)))
            data.Columns.Add(New DataColumn("object", GetType(Object)))
            data.Columns.Add(New DataColumn("datetime", GetType(DateTime)))
            data.Columns.Add(New DataColumn("double", GetType(Double)))
            Dim row As DataRow
            row = data.NewRow
            row("some values") = "this is a Test!"
            row("string") = "a nice string, isn't it?"
            row("int16") = Int16.MinValue
            row("int32") = Int32.MinValue
            row("int64") = Int64.MinValue 'not supported, only int32!
            row("boolean") = False
            row("object") = New Object
            row("datetime") = DateTime.MaxValue
            row("double") = Double.MinValue
            data.Rows.Add(row)
            row = data.NewRow
            row("string") = "=""this should not be interpreted as a formula"""
            row("int16") = Int16.MaxValue
            row("int32") = Int32.MaxValue
            row("int64") = Int64.MaxValue 'not supported, only int32!
            row("boolean") = True
            row("object") = New DateTime
            row("datetime") = DateTime.MaxValue
            row("double") = Double.MaxValue
            data.Rows.Add(row)
            row = data.NewRow
            row("some values") = DBNull.Value
            row("string") = Nothing
            row("int16") = Int16.Parse("0")
            row("int32") = Int32.Parse("0")
            row("int64") = Int64.Parse("0")
            row("boolean") = False
            row("object") = DBNull.Value
            row("datetime") = New DateTime(2005, 9, 29, 13, 50, 20, 997)
            row("double") = DBNull.Value
            data.Rows.Add(row)
            row = data.NewRow
            row("some values") = "''apostrophes'"
            row("string") = Nothing
            row("int16") = Int16.Parse("1000")
            row("int32") = Int32.Parse("10000000")
            row("int64") = Int64.Parse("-10000000")
            row("boolean") = False
            row("object") = DBNull.Value
            row("datetime") = New DateTime(1975, 9, 29, 13, 50, 20, 997)
            row("double") = Double.Parse("0")
            data.Rows.Add(row)
            'Write test data
            CompuMaster.Data.XlsEpplus.WriteDataTableToXlsFile(Nothing, TempFile, data, "test")
            data.TableName = "Data as written to file"
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(data))

            'Read and compare written test data
            '==================================

            'read the existing file, auto-detect column-types, take datatable and compare it with the written data: it should be always the same (or must be argumented and discussed with Jochen why it isn't)
            'the number of columns and rows should be always 2
            Dim ReReadData As DataTable
            ReReadData = CompuMaster.Data.XlsEpplus.ReadDataTableFromXlsFile(TempFile, "test", True)
            Assert.AreEqual("test", ReReadData.TableName, "SaveAndReadDataTypes #05")
            ReReadData.TableName = "Data as read from file"
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(ReReadData))
            Assert.AreEqual(9, ReReadData.Columns.Count, "SaveAndReadDataTypes #06")
            Assert.AreEqual("some values", ReReadData.Columns(0).ColumnName, "SaveAndReadDataTypes #07")
            Assert.AreEqual(4, ReReadData.Rows.Count, "SaveAndReadDataTypes #08")
            'column data types
            Assert.AreEqual(GetType(String), ReReadData.Columns(0).DataType, "SaveAndReadDataTypes #11")
            Assert.AreEqual(GetType(String), ReReadData.Columns(1).DataType, "SaveAndReadDataTypes #12")
            Assert.AreEqual(GetType(Double), ReReadData.Columns(2).DataType, "SaveAndReadDataTypes #13")
            Assert.AreEqual(GetType(Double), ReReadData.Columns(3).DataType, "SaveAndReadDataTypes #14")
            Assert.AreEqual(GetType(Double), ReReadData.Columns(4).DataType, "SaveAndReadDataTypes #15")
            Assert.AreEqual(GetType(Boolean), ReReadData.Columns(5).DataType, "SaveAndReadDataTypes #16")
            Assert.AreEqual(GetType(String), ReReadData.Columns(6).DataType, "SaveAndReadDataTypes #17")
            Assert.AreEqual(GetType(DateTime), ReReadData.Columns(7).DataType, "SaveAndReadDataTypes #18")
            Assert.AreEqual(GetType(Double), ReReadData.Columns(8).DataType, "SaveAndReadDataTypes #19")
            'row 1
            Assert.AreEqual("this is a Test!", ReReadData.Rows(0)(0), "SaveAndReadDataTypes #21")
            Assert.AreEqual("a nice string, isn't it?", ReReadData.Rows(0)(1), "SaveAndReadDataTypes #22")
            Assert.AreEqual(Int16.MinValue, ReReadData.Rows(0)(2), "SaveAndReadDataTypes #23")
            Assert.AreEqual(Int32.MinValue, ReReadData.Rows(0)(3), "SaveAndReadDataTypes #24")
            Assert.AreEqual(Int64.MinValue, ReReadData.Rows(0)(4), "SaveAndReadDataTypes #25")
            Assert.AreEqual(False, ReReadData.Rows(0)(5), "SaveAndReadDataTypes #26")
            Assert.AreEqual(New Object().ToString, ReReadData.Rows(0)(6), "SaveAndReadDataTypes #27")
            Assert.AreEqual(New DateTime(9999, 12, 31, 23, 59, 59, 999), ReReadData.Rows(0)(7), "SaveAndReadDataTypes #28")
            Assert.AreEqual(Double.MinValue, ReReadData.Rows(0)(8), "SaveAndReadDataTypes #29")
            'row 2
            Assert.AreEqual(DBNull.Value, ReReadData.Rows(1)(0), "SaveAndReadDataTypes #31")
            Assert.AreEqual("=""this should not be interpreted as a formula""", ReReadData.Rows(1)(1), "SaveAndReadDataTypes #32")
            Assert.AreEqual(Int16.MaxValue, ReReadData.Rows(1)(2), "SaveAndReadDataTypes #33")
            Assert.AreEqual(Int32.MaxValue, ReReadData.Rows(1)(3), "SaveAndReadDataTypes #34")
            Assert.AreEqual(Int64.MaxValue, ReReadData.Rows(1)(4), "SaveAndReadDataTypes #35")
            Assert.AreEqual(True, ReReadData.Rows(1)(5), "SaveAndReadDataTypes #36")
            Assert.AreEqual(DBNull.Value, ReReadData.Rows(1)(6), "SaveAndReadDataTypes #37")
            Assert.AreEqual(New DateTime(9999, 12, 31, 23, 59, 59, 999), ReReadData.Rows(1)(7), "SaveAndReadDataTypes #38")
            Assert.AreEqual(Double.MaxValue, ReReadData.Rows(1)(8), "SaveAndReadDataTypes #39")
            'row 3
            Assert.AreEqual(DBNull.Value, ReReadData.Rows(2)(0), "SaveAndReadDataTypes #41")
            Assert.AreEqual(DBNull.Value, ReReadData.Rows(2)(1), "SaveAndReadDataTypes #42")
            Assert.AreEqual(Int16.Parse("0"), ReReadData.Rows(2)(2), "SaveAndReadDataTypes #43")
            Assert.AreEqual(Int32.Parse("0"), ReReadData.Rows(2)(3), "SaveAndReadDataTypes #44")
            Assert.AreEqual(Int64.Parse("0"), ReReadData.Rows(2)(4), "SaveAndReadDataTypes #45")
            Assert.AreEqual(False, ReReadData.Rows(2)(5), "SaveAndReadDataTypes #46")
            Assert.AreEqual(DBNull.Value, ReReadData.Rows(2)(6), "SaveAndReadDataTypes #47")
            Assert.AreEqual(New DateTime(2005, 9, 29, 13, 50, 20, 997), ReReadData.Rows(2)(7), "SaveAndReadDataTypes #48")
            Assert.AreEqual(DBNull.Value, ReReadData.Rows(2)(8), "SaveAndReadDataTypes #49")
            'row 4
            Assert.AreEqual("''apostrophes'", ReReadData.Rows(3)(0), "SaveAndReadDataTypes #51")
            Assert.AreEqual(DBNull.Value, ReReadData.Rows(3)(1), "SaveAndReadDataTypes #52")
            Assert.AreEqual(Int16.Parse("1000"), ReReadData.Rows(3)(2), "SaveAndReadDataTypes #53")
            Assert.AreEqual(Int32.Parse("10000000"), ReReadData.Rows(3)(3), "SaveAndReadDataTypes #54")
            Assert.AreEqual(Int64.Parse("-10000000"), ReReadData.Rows(3)(4), "SaveAndReadDataTypes #55")
            Assert.AreEqual(False, ReReadData.Rows(3)(5), "SaveAndReadDataTypes #56")
            Assert.AreEqual(DBNull.Value, ReReadData.Rows(3)(6), "SaveAndReadDataTypes #57")
            Assert.AreEqual(New DateTime(1975, 9, 29, 13, 50, 20, 997), ReReadData.Rows(3)(7), "SaveAndReadDataTypes #58")
            Assert.AreEqual(Double.Parse("0"), ReReadData.Rows(3)(8), "SaveAndReadDataTypes #59")
        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <remarks>
        ''' write 5 different column data types
        ''' create a table structure for reading the 1st, 3rd and 5th column
        ''' read the data (only the defined 3 columns should be imported since column names in excel sheet (see content of first row!) and column names in datatable match respectively when firstRowContainsColumnNames is true then the datatable's column names must match the column index in excel ("1", "2", "3", ...))
        ''' compare/validate the data
        ''' </remarks>
        ''' <history>
        ''' 	[adminwezel]	02.02.2007	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        <Test()> Public Sub SaveFiveColumnsAndReadFirstAndThirdAndFifthColumn()
            'Prepare test data
            Dim data As New DataTable("testtable")
            data.Columns.Add(New DataColumn("1", GetType(Int32)))
            data.Columns.Add(New DataColumn("2", GetType(Int32)))
            data.Columns.Add(New DataColumn("3", GetType(Int32)))
            data.Columns.Add(New DataColumn("4", GetType(Int32)))
            data.Columns.Add(New DataColumn("5", GetType(Int32)))
            Dim row As DataRow = data.NewRow
            row("1") = 6
            row("2") = 7
            row("3") = 8
            row("4") = 9
            row("5") = 10
            data.Rows.Add(row)
            'Write test data
            CompuMaster.Data.XlsEpplus.WriteDataTableToXlsFile(Nothing, TempFile, data, "test")

            'Read and compare written test data
            '==================================

            'read the existing file, auto-detect column-types, take datatable and compare it with the written data: it should be always the same (or must be argumented and discussed with Jochen why it isn't)
            'the number of columns and rows should be always 2
            Dim ReReadData As DataTable
            ReReadData = CompuMaster.Data.XlsEpplus.ReadDataTableFromXlsFile(TempFile, "test", True)
            Assert.AreEqual("test", ReReadData.TableName, "SaveFiveColumnsAndReadFirstAndThirdAndFifthColumn #05")
            Assert.AreEqual(5, ReReadData.Columns.Count, "SaveFiveColumnsAndReadFirstAndThirdAndFifthColumn #10")
            Assert.AreEqual("1", ReReadData.Columns(0).ColumnName, "SaveFiveColumnsAndReadFirstAndThirdAndFifthColumn #11")
            Assert.AreEqual("2", ReReadData.Columns(1).ColumnName, "SaveFiveColumnsAndReadFirstAndThirdAndFifthColumn #12")
            Assert.AreEqual("3", ReReadData.Columns(2).ColumnName, "SaveFiveColumnsAndReadFirstAndThirdAndFifthColumn #13")
            Assert.AreEqual("4", ReReadData.Columns(3).ColumnName, "SaveFiveColumnsAndReadFirstAndThirdAndFifthColumn #14")
            Assert.AreEqual("5", ReReadData.Columns(4).ColumnName, "SaveFiveColumnsAndReadFirstAndThirdAndFifthColumn #15")
            Assert.AreEqual(1, ReReadData.Rows.Count, "SaveFiveColumnsAndReadFirstAndThirdAndFifthColumn #20")
            Assert.AreEqual(6, ReReadData.Rows(0)("1"), "SaveFiveColumnsAndReadFirstAndThirdAndFifthColumn #21")
            Assert.AreEqual(7, ReReadData.Rows(0)("2"), "SaveFiveColumnsAndReadFirstAndThirdAndFifthColumn #22")
            Assert.AreEqual(8, ReReadData.Rows(0)("3"), "SaveFiveColumnsAndReadFirstAndThirdAndFifthColumn #23")
            Assert.AreEqual(9, ReReadData.Rows(0)("4"), "SaveFiveColumnsAndReadFirstAndThirdAndFifthColumn #24")
            Assert.AreEqual(10, ReReadData.Rows(0)("5"), "SaveFiveColumnsAndReadFirstAndThirdAndFifthColumn #25")

        End Sub

        <Test> Public Sub ReadEmptySheetDimensions()
            Dim file As String = GlobalTestSetup.PathToTestFiles("testfiles\emptysheets.xlsx")
            Dim ds As DataSet = CompuMaster.Data.XlsEpplus.ReadDataSetFromXlsFile(file, False)
            Assert.Pass()
        End Sub

        <Test()> Public Sub ReadDataTypes()
            Dim file As String = GlobalTestSetup.PathToTestFiles("testfiles\datatype-checks.xlsx")
            Dim dt As DataTable = CompuMaster.Data.XlsEpplus.ReadDataTableFromXlsFile(file, True)
            Console.WriteLine(ColumnDataTypesToPlainTextTableFixedColumnWidths(dt))
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt))
            Assert.AreEqual(27, dt.Columns.Count, "Col-Length")
            Assert.AreEqual(1, dt.Rows.Count, "Row-Length")
            Assert.AreEqual(True, dt.Rows(0)(0), "Boolean")
            Assert.AreEqual(New DateTime(1944, 12, 1), dt.Rows(0)(1), "1.12.1944")
            Assert.AreEqual(New DateTime(2144, 12, 1), dt.Rows(0)(2), "1.12.2144")
            If CType(dt.Rows(0)(3), DateTime) = New DateTime(1900, 1, 1) Then
                'Epplus with fixed bug
                Assert.AreEqual(New DateTime(1900, 1, 1), dt.Rows(0)(3), "1.1.1900")
            Else
                ''Epplus with bug (prooven with v4.5.3.2), see https://github.com/JanKallman/EPPlus/issues/574
                Assert.AreEqual(New DateTime(1899, 12, 31), dt.Rows(0)(3), "1.1.1900 incorrectly converted by Epplus to 31.12.1899 (bug at Epplus)")
            End If
            Assert.AreEqual(New DateTime(9999, 12, 31), dt.Rows(0)(4), "1.1.1900")
            Assert.AreEqual("01.01.1600", dt.Rows(0)(5), "1.1.1600 always < 1900 so excel handles it as string")
            Assert.AreEqual(New DateTime(1905, 1, 1, 15, 15, 0), dt.Rows(0)(6), "1.1.1905 15:15:00")
            Assert.AreEqual(Double.NaN, dt.Rows(0)(7), "#DIV/0")
            Assert.AreEqual("#NAME?", dt.Rows(0)(8), "#NAME?")
            Assert.AreEqual("#VALUE!", dt.Rows(0)(9), "#WERT")
            Assert.AreEqual("#REF!", dt.Rows(0)(10), "#BEZUG")
            Assert.AreEqual(0, dt.Rows(0)(11), "circular reference partner cell #1")
            Assert.AreEqual(0, dt.Rows(0)(12), "circular reference partner cell #2")
            Assert.AreEqual("test", dt.Rows(0)(13), "test")
            Assert.AreEqual("test", dt.Rows(0)(14), "'test")
            Assert.AreEqual("'test", dt.Rows(0)(15), "''test")
            'Check time values
            Dim BaseDateForAllTimeValues As DateTime
            If CType(dt.Rows(0)(3), DateTime) = New DateTime(1900, 1, 1) Then
                'Epplus with fixed bug
                BaseDateForAllTimeValues = New DateTime(1900, 1, 1)
            Else
                'Epplus with bug (prooven with v4.5.3.2), see https://github.com/JanKallman/EPPlus/issues/574
                BaseDateForAllTimeValues = New DateTime(1899, 12, 31)
            End If
            BaseDateForAllTimeValues = BaseDateForAllTimeValues.AddDays(-1) 'All time-only values start 1 day before 1900/01/01
            Assert.AreEqual(BaseDateForAllTimeValues.Add(New TimeSpan(15, 15, 0)), dt.Rows(0)(16), "15:15:00")
            Assert.AreEqual(BaseDateForAllTimeValues.Add(New TimeSpan(15, 35, 34)), dt.Rows(0)(17), "15:35:34")
            Assert.AreEqual(BaseDateForAllTimeValues.Add(New TimeSpan(256, 25, 20)), dt.Rows(0)(18), "256:25:20 alias excel-internal 10.01.1900 16:25:20")
            Assert.AreEqual(BaseDateForAllTimeValues.Add(New TimeSpan(256, 25, 20)), dt.Rows(0)(19), "256:25:20 alias excel-internal 10.01.1900 16:25:20")
            Assert.AreEqual(2.0, dt.Rows(0)(20), "Byte 2")
            Assert.AreEqual(1.45325, dt.Rows(0)(21), "Single 1,45325")
            Assert.AreEqual("D", dt.Rows(0)(22), "Char D")
            Assert.AreEqual(39211212.3434733, dt.Rows(0)(23), "Decimal 39211212,3434733")
            Assert.AreEqual(289382.0, dt.Rows(0)(24), "Integer 289382")
            Assert.AreEqual(2297987128.0, dt.Rows(0)(25), "Long 2297987128")
            Assert.AreEqual(312.0, dt.Rows(0)(26), "Short 312")
        End Sub

        <Test()> Public Sub ReadErrorTypes()
            Dim file As String = GlobalTestSetup.PathToTestFiles("testfiles\errortype-checks.xlsx")
            Dim dt As DataTable = CompuMaster.Data.XlsEpplus.ReadDataTableFromXlsFile(file, True)
            Console.WriteLine(ColumnDataTypesToPlainTextTableFixedColumnWidths(dt))
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt))
            Assert.AreEqual(19, dt.Columns.Count, "Col-Length")
            Assert.AreEqual(7, dt.Rows.Count, "Row-Length")
            For MyCounter As Integer = 0 To 11
                Select Case MyCounter
                    Case 0, 5, 6, 11
                        Assert.AreEqual(GetType(Double), dt.Columns(MyCounter).DataType, "cols index " & MyCounter & " should be number field, but Strings due to Error values")
                    Case Else
                        Assert.AreEqual(GetType(String), dt.Columns(MyCounter).DataType, "cols index " & MyCounter & " should be string field due to Error values")
                End Select
            Next
            For MyCounter As Integer = 12 To 18
                Select Case MyCounter
                    Case 12, 17
                        Assert.AreEqual(GetType(Double), dt.Columns(MyCounter).DataType, "cols index " & MyCounter & " should be number field, but Strings due to Error values")
                    Case 18
                        Assert.AreEqual(GetType(String), dt.Columns(MyCounter).DataType, "all datatypes of err-only-cols can't be identified --> string")
                        Assert.AreEqual(True, CType(dt.Rows(1)(MyCounter), String).StartsWith("#"), "all datatypes of err-only-cols can't be identified --> string")
                        Assert.AreEqual(True, CType(dt.Rows(2)(MyCounter), String).StartsWith("#"), "all datatypes of err-only-cols can't be identified --> string")
                        Assert.AreEqual(True, CType(dt.Rows(3)(MyCounter), String).StartsWith("#"), "all datatypes of err-only-cols can't be identified --> string")
                        Assert.AreEqual(True, CType(dt.Rows(4)(MyCounter), String).StartsWith("#"), "all datatypes of err-only-cols can't be identified --> string")
                        Assert.AreEqual(True, CType(dt.Rows(5)(MyCounter), String).StartsWith("#"), "all datatypes of err-only-cols can't be identified --> string")
                        Assert.AreEqual(True, CType(dt.Rows(6)(MyCounter), String).StartsWith("#"), "all datatypes of err-only-cols can't be identified --> string")
                    Case Else
                        Assert.AreEqual(GetType(String), dt.Columns(MyCounter).DataType, "all datatypes of err-only-cols can't be identified --> string")
                        Assert.AreEqual(True, CType(dt.Rows(0)(MyCounter), String).StartsWith("#"), "all datatypes of err-only-cols can't be identified --> string")
                        Assert.AreEqual(True, CType(dt.Rows(1)(MyCounter), String).StartsWith("#"), "all datatypes of err-only-cols can't be identified --> string")
                End Select
            Next
            Assert.AreEqual(1, dt.Rows(1)(0), "static 1 as number")
            Assert.AreEqual("1", dt.Rows(1)(1), "Static 1 As number")
            Assert.AreEqual("1", dt.Rows(1)(2), "static 1 as number")
            Assert.AreEqual("1", dt.Rows(1)(3), "Static 1 As number")
            Assert.AreEqual("1", dt.Rows(1)(4), "static 1 as number")
            Assert.AreEqual(1, dt.Rows(1)(5), "static 1 as number")
            Assert.AreEqual(1, dt.Rows(0)(6), "static 1 as number")
            Assert.AreEqual("1", dt.Rows(0)(7), "static 1 as number")
            Assert.AreEqual("1", dt.Rows(0)(8), "static 1 as number")
            Assert.AreEqual("1", dt.Rows(0)(9), "static 1 as number")
            Assert.AreEqual("1", dt.Rows(0)(10), "static 1 as number")
            Assert.AreEqual(1, dt.Rows(0)(11), "static 1 as number")
            'TODO: compare expected results with MSExcelOleDbProvider
            '1st or 2nd line cells with errors
            Assert.AreEqual(Double.NaN, dt.Rows(0)(0), "#DIV/0")
            Assert.AreEqual("#NAME?", dt.Rows(0)(1), "#NAME?")
            Assert.AreEqual("#VALUE!", dt.Rows(0)(2), "#WERT")
            Assert.AreEqual("#REF!", dt.Rows(0)(3), "#BEZUG!")
            Assert.AreEqual("#N/A", dt.Rows(0)(4), "#NV!")
            Assert.AreEqual(Double.PositiveInfinity, dt.Rows(0)(5), "#ZAHL!")
            Assert.AreEqual(Double.NaN, dt.Rows(1)(6), "#DIV/0")
            Assert.AreEqual("#NAME?", dt.Rows(1)(7), "#NAME?")
            Assert.AreEqual("#VALUE!", dt.Rows(1)(8), "#WERT")
            Assert.AreEqual("#REF!", dt.Rows(1)(9), "#BEZUG!")
            Assert.AreEqual("#N/A", dt.Rows(1)(10), "#NV!")
            Assert.AreEqual(Double.PositiveInfinity, dt.Rows(1)(11), "#ZAHL!")
            'All lines and cells with errors
            Assert.AreEqual(Double.NaN, dt.Rows(0)(12), "#DIV/0")
            Assert.AreEqual("#NAME?", dt.Rows(0)(13), "#NAME?")
            Assert.AreEqual("#VALUE!", dt.Rows(0)(14), "#WERT")
            Assert.AreEqual("#REF!", dt.Rows(0)(15), "#BEZUG!")
            Assert.AreEqual("#N/A", dt.Rows(0)(16), "#NV!")
            Assert.AreEqual(Double.PositiveInfinity, dt.Rows(0)(17), "#ZAHL!")
            Assert.AreEqual(Double.NaN, dt.Rows(1)(12), "#DIV/0")
            Assert.AreEqual("#NAME?", dt.Rows(1)(13), "#NAME?")
            Assert.AreEqual("#VALUE!", dt.Rows(1)(14), "#WERT")
            Assert.AreEqual("#REF!", dt.Rows(1)(15), "#BEZUG!")
            Assert.AreEqual("#N/A", dt.Rows(1)(16), "#NV!")
            Assert.AreEqual(Double.PositiveInfinity, dt.Rows(1)(17), "#ZAHL!")
            'Column with allmost all cells with errors
            Assert.AreEqual("1", dt.Rows(0)(18), "#DIV/0")
            Assert.AreEqual("#DIV/0!", dt.Rows(1)(18), "#DIV/0")
            Assert.AreEqual("#NAME?", dt.Rows(2)(18), "#NAME?")
            Assert.AreEqual("#VALUE!", dt.Rows(3)(18), "#WERT")
            Assert.AreEqual("#REF!", dt.Rows(4)(18), "#BEZUG!")
            Assert.AreEqual("#N/A", dt.Rows(5)(18), "#NV!")
            Assert.AreEqual("#NUM!", dt.Rows(6)(18), "#ZAHL!")
        End Sub

        ''' <summary>
        ''' Provide information on table schema
        ''' </summary>
        ''' <param name="table"></param>
        ''' <returns></returns>
        Private Function ColumnDataTypesToPlainTextTableFixedColumnWidths(table As DataTable) As String
            Dim dt As New DataTable(table.TableName & " - TableSchema")
            For MyCounter As Integer = 0 To table.Columns.Count - 1
                dt.Columns.Add(table.Columns(MyCounter).ColumnName, GetType(String))
            Next
            Dim row As DataRow = dt.NewRow
            For MyCounter As Integer = 0 To dt.Columns.Count - 1
                row(MyCounter) = table.Columns(MyCounter).DataType.ToString
            Next
            dt.Rows.Add(row)
            Return CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt)
        End Function

        <Test()> Public Sub ReadTestFileVIProjektFixedStringColTypes()
            Dim file As String = GlobalTestSetup.PathToTestFiles("testfiles\vi-projekte.xlsx")
            Dim dt As New DataTable("Root")
            dt.Columns.Add("ADM", GetType(String))
            dt.Columns.Add("ProdFam", GetType(String))
            dt.Columns.Add("Stück", GetType(String))
            dt.Columns.Add("Kunde", GetType(String))
            dt.Columns.Add("Modell", GetType(String))
            dt.Columns.Add("Limitpreis", GetType(String))
            dt.Columns.Add("Planmonat", GetType(String))
            dt.Columns.Add("Chance", GetType(String))
            dt.Columns.Add("Umsatzgewicht", GetType(String))
            dt.Columns.Add("Lost Order", GetType(String))
            dt.Columns.Add("Anmerkungen", GetType(String))
            CompuMaster.Data.XlsEpplus.ReadDataTableFromXlsFile(file, "Alban", False, dt)
            Assert.AreEqual(11, dt.Columns.Count, "Col-Length")
            Assert.AreEqual(19, dt.Rows.Count, "Row-Length")
        End Sub

        <Test()> Public Sub ReadTestFileVIProjektFixedColTypes()
            Dim file As String = GlobalTestSetup.PathToTestFiles("testfiles\vi-projekte.xlsx")
            Dim dt As New DataTable("Root")
            dt.Columns.Add("ADM", GetType(String))
            dt.Columns.Add("ProdFam", GetType(Double))
            dt.Columns.Add("Stück", GetType(Double))
            dt.Columns.Add("Kunde", GetType(String))
            dt.Columns.Add("Modell", GetType(String))
            dt.Columns.Add("Limitpreis", GetType(String))
            dt.Columns.Add("Planmonat", GetType(String))
            dt.Columns.Add("Chance", GetType(String))
            dt.Columns.Add("Umsatzgewicht", GetType(String))
            dt.Columns.Add("Lost Order", GetType(String))
            dt.Columns.Add("Anmerkungen", GetType(String))
            CompuMaster.Data.XlsEpplus.ReadDataTableFromXlsFile(file, "Alban", True, dt)
        End Sub

        <Test()> Public Sub ReadTestFileVIProjektDynamicColTypes()
            Dim file As String = GlobalTestSetup.PathToTestFiles("testfiles\vi-projekte.xlsx")
            Dim dt As DataTable = CompuMaster.Data.XlsEpplus.ReadDataTableFromXlsFile(file, "Alban", True)
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt))
            Assert.AreEqual(GetType(String), dt.Columns("Info").DataType)
            Assert.AreEqual(GetType(Double), dt.Columns("KZ").DataType)
            Assert.AreEqual(GetType(Double), dt.Columns("Stück").DataType)
            Assert.AreEqual(GetType(String), dt.Columns("Buchung").DataType)
            Assert.AreEqual(GetType(String), dt.Columns("Modell").DataType)
            Assert.AreEqual(GetType(Double), dt.Columns("Rabatt").DataType)
            Assert.AreEqual(GetType(String), dt.Columns("Planmonat").DataType)
            Assert.AreEqual(GetType(Double), dt.Columns("Chance").DataType)
            Assert.AreEqual(GetType(Double), dt.Columns("Gewicht").DataType)
            Assert.AreEqual(GetType(String), dt.Columns("Lost Order").DataType)
            Assert.AreEqual(GetType(String), dt.Columns("Anmerkungen").DataType)
        End Sub

        <Test()> Public Sub ReadTestFileVIProjektEndOfContentRowColIndexes()
            Dim file As String = GlobalTestSetup.PathToTestFiles("testfiles\vi-projekte.xlsx")
            Dim dt As DataTable = CompuMaster.Data.XlsEpplus.ReadDataTableFromXlsFile(file, "Alban", True)
            Assert.AreEqual(11, dt.Columns.Count, "Col-Length")
            Assert.AreEqual(18, dt.Rows.Count, "Row-Length")
        End Sub

        <Test()> Public Sub ReadTestFileQnA()
            Dim file As String = GlobalTestSetup.PathToTestFiles("testfiles\QuestsNAnswers.xlsx")
            Dim ds As DataSet = CompuMaster.Data.XlsEpplus.ReadDataSetFromXlsFile(file, True)
            Assert.AreEqual(45, ds.Tables(0).Rows.Count, "Row-Length")
            Assert.AreEqual(35, ds.Tables(1).Rows.Count, "Row-Length")
        End Sub

    End Class

End Namespace