Imports NUnit.Framework

Namespace CompuMaster.Test.Data

    <Parallelizable(ParallelScope.All)>
    <TestFixture>
    Public Class WorksheetsIndexBase1Or0Test

        <Test> Public Sub GlobalFirstWorksheetIndex()
#If NETFRAMEWORK Then
            Assert.AreEqual(1, CompuMaster.Data.XlsEpplus.GlobalFirstWorksheetBaseIndex)
#Else
            Assert.AreEqual(0, CompuMaster.Data.XlsEpplus.GlobalFirstWorksheetBaseIndex)
#End If
        End Sub

    End Class

End Namespace