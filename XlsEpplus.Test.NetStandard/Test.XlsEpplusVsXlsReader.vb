Imports NUnit.Framework
Imports System.Data

Namespace CompuMaster.Test.Data

    <TestFixture()> Public Class XlsEpplusVsXlsReader
        Implements IDisposable

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

#If Not CI_Build Then
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
            Dim ReReadData As DataTable = Nothing
            Try
                ReReadData = CompuMaster.Data.XlsReader.ReadDataTableFromXlsFile(TempFile, "test")
            Catch ex As Exception
                Assert.Ignore("MS Excel Provider support not installed for current platform " & System.Environment.OSVersion.Platform & "/" & PlatformDependentProcessBitNumber() & " (" & System.Environment.OSVersion.ToString & ")")
            End Try
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
            Dim ReReadData As DataTable = Nothing
            Try
                ReReadData = CompuMaster.Data.XlsReader.ReadDataTableFromXlsFile(TempFile, "test")
            Catch ex As Exception
                Assert.Ignore("MS Excel Provider support not installed for current platform " & System.Environment.OSVersion.Platform & "/" & PlatformDependentProcessBitNumber() & " (" & System.Environment.OSVersion.ToString & ")")
            End Try
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
            Dim ReReadData As DataTable = Nothing
            Try
                ReReadData = CompuMaster.Data.XlsReader.ReadDataTableFromXlsFile(TempFile, "test")
            Catch ex As Exception
                Assert.Ignore("MS Excel Provider support not installed for current platform " & System.Environment.OSVersion.Platform & "/" & PlatformDependentProcessBitNumber() & " (" & System.Environment.OSVersion.ToString & ")")
            End Try
            Assert.AreEqual("line 1" & System.Environment.NewLine & "line 2" & System.Environment.NewLine & "line 3" & System.Environment.NewLine & "line 4", ReReadData.Rows(0)(0), "SaveAndReadExtraLargeFieldsWithLineBreaks #11")
        End Sub

        <Test()> Public Sub SaveAndReadLastCell()
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
            Dim ReReadData As DataTable = Nothing
            Try
                ReReadData = CompuMaster.Data.XlsReader.ReadDataTableFromXlsFile(TempFile, "test")
            Catch ex As Exception
                Assert.Ignore("MS Excel Provider support not installed for current platform " & System.Environment.OSVersion.Platform & "/" & PlatformDependentProcessBitNumber() & " (" & System.Environment.OSVersion.ToString & ")")
            End Try
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
            Dim ReReadData As DataTable = Nothing
            Try
                ReReadData = CompuMaster.Data.XlsReader.ReadDataTableFromXlsFile(TempFile, "test")
            Catch ex As Exception
                Assert.Ignore("MS Excel Provider support not installed for current platform " & System.Environment.OSVersion.Platform & "/" & PlatformDependentProcessBitNumber() & " (" & System.Environment.OSVersion.ToString & ")")
            End Try
            Assert.AreEqual(4, ReReadData.Rows.Count, "SaveAndReadEmptyStates #10") 'because last 2 lines only contains DBNull
            Assert.AreEqual(DBNull.Value, ReReadData.Rows(0)(0), "SaveAndReadEmptyStates #11")
            Assert.AreEqual("", ReReadData.Rows(1)(0), "SaveAndReadEmptyStates #12")
            Assert.AreEqual(DBNull.Value, ReReadData.Rows(2)(0), "SaveAndReadEmptyStates #13")
            Assert.AreEqual(DBNull.Value, ReReadData.Rows(3)(0), "SaveAndReadEmptyStates #14")
        End Sub

        Private Function PlatformDependentProcessBitNumber() As String
            If Environment.Is64BitProcess Then
                Return "x64"
            Else
                Return "x32"
            End If
        End Function

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
            Dim ReReadData As DataTable = Nothing
            Try
                ReReadData = CompuMaster.Data.XlsReader.ReadDataTableFromXlsFile(TempFile, "test")
            Catch ex As Exception
                Assert.Ignore("MS Excel Provider support not installed for current platform " & System.Environment.OSVersion.Platform & "/" & PlatformDependentProcessBitNumber() & " (" & System.Environment.OSVersion.ToString & ")")
            End Try
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
#End If

        Private disposedValue As Boolean = False        ' So ermitteln Sie überflüssige Aufrufe

        ''' <summary>
        ''' Clean up of temp file
        ''' </summary>
        ''' <param name="disposing"></param>
        ''' <remarks></remarks>
        Protected Overridable Sub Dispose(ByVal disposing As Boolean)
            If Not Me.disposedValue Then
                If System.IO.File.Exists(Me.TempFile) Then
                    Try
                        System.IO.File.Delete(Me.TempFile)
#Disable Warning CA1031 ' Do not catch general exception types
                    Catch
#Enable Warning CA1031 ' Do not catch general exception types
                    End Try
                End If
            End If
            Me.disposedValue = True
        End Sub

#Region " IDisposable Support "
        ' Dieser Code wird von Visual Basic hinzugefügt, um das Dispose-Muster richtig zu implementieren.
        Public Sub Dispose() Implements IDisposable.Dispose
            ' Ändern Sie diesen Code nicht. Fügen Sie oben in Dispose(ByVal disposing As Boolean) Bereinigungscode ein.
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub
#End Region

    End Class

End Namespace