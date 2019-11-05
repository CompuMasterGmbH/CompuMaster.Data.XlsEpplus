﻿Namespace CompuMaster.Test.Data

    Public Class GlobalTestSetup

        Public Shared Function PathToTestFiles(subPath As String) As String
            Return System.IO.Path.Combine(System.IO.Path.Combine(System.Reflection.Assembly.GetExecutingAssembly.Location, ".."), subPath.Replace("\", System.IO.Path.DirectorySeparatorChar))
        End Function

    End Class

End Namespace