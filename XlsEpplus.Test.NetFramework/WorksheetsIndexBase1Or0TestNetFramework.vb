Imports NUnit.Framework

Namespace CompuMaster.Test.Data

    <TestFixture> Public Class WorksheetsIndexBase1Or0Test

        <Test> Public Sub GlobalFirstWorksheetIndex()
            Assert.AreEqual(1, CompuMaster.Data.XlsEpplus.GlobalFirstWorksheetBaseIndex)
        End Sub

    End Class

End Namespace