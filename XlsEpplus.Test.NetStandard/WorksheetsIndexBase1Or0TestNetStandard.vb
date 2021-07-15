Imports NUnit.Framework

Namespace CompuMaster.Test.Data

    <TestFixture> Public Class WorksheetsIndexBase1Or0TestNetStandard

        <Test> Public Sub GlobalFirstWorksheetIndex()
            Assert.AreEqual(0, CompuMaster.Data.XlsEpplus.GlobalFirstWorksheetBaseIndex)
        End Sub

    End Class

End Namespace