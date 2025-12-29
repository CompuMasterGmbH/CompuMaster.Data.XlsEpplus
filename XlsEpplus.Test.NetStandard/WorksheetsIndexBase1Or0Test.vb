Imports NUnit.Framework
Imports NUnit.Framework.Legacy

Namespace CompuMaster.Test.Data

    <Parallelizable(ParallelScope.All)>
    <TestFixture>
    Public Class WorksheetsIndexBase1Or0Test

        <AutoIgnoreOnNonWindowsNativeLoadFailure>
        <Test> Public Sub GlobalFirstWorksheetIndex()
#If NETFRAMEWORK Then
            ClassicAssert.AreEqual(1, CompuMaster.Data.XlsEpplus.GlobalFirstWorksheetBaseIndex)
#Else
            ClassicAssert.AreEqual(0, CompuMaster.Data.XlsEpplus.GlobalFirstWorksheetBaseIndex)
#End If
        End Sub

    End Class

End Namespace