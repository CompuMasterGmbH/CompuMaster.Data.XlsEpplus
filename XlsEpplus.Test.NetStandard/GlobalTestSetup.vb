Imports System.Reflection
Imports System.Runtime.InteropServices
Imports CompuMaster.Test.Data
Imports NUnit.Framework
Imports NUnit.Framework.Interfaces
Imports NUnit.Framework.Internal.Commands

<Assembly: AutoIgnoreOnNonWindowsNativeLoadFailure()>

Namespace CompuMaster.Test.Data

    Public NotInheritable Class GlobalTestSetup

        Public Shared Function PathToTestFiles(subPath As String) As String
            'Originates from .NET Framework version - might be dropped if all tests successful
            'Return System.IO.Path.Combine(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly.Location), subPath.Replace("\", System.IO.Path.DirectorySeparatorChar))
            Return System.IO.Path.Combine(System.IO.Path.Combine(System.Reflection.Assembly.GetExecutingAssembly.Location, ".."), subPath.Replace("\", System.IO.Path.DirectorySeparatorChar))
        End Function

    End Class

    <AttributeUsage(AttributeTargets.Assembly Or AttributeTargets.Class Or AttributeTargets.Method, AllowMultiple:=False, Inherited:=True)>
    Public NotInheritable Class AutoIgnoreOnNonWindowsNativeLoadFailureAttribute
        Inherits NUnitAttribute
        Implements IWrapTestMethod

        Public Function Wrap(command As TestCommand) As TestCommand Implements IWrapTestMethod.Wrap
            Return New WrappedCommand(command)
        End Function

        Private NotInheritable Class WrappedCommand
            Inherits TestCommand

            Private ReadOnly _inner As TestCommand

            Public Sub New(innerCommand As TestCommand)
                MyBase.New(innerCommand.Test)
                _inner = innerCommand
            End Sub

            Public Overrides Function Execute(context As NUnit.Framework.Internal.TestExecutionContext) As NUnit.Framework.Internal.TestResult
                Try
                    Return _inner.Execute(context)
                Catch ex As Exception When ShouldIgnore(ex)
                    Assert.Ignore("Auto-ignored on non-Windows (native dependency missing or library support too limited in System.Drawing on non-windows platforms): " & ex.GetType().Name)
                    Throw
                End Try
            End Function

            Private Shared Function ShouldIgnore(ex As Exception) As Boolean
                If RuntimeInformation.IsOSPlatform(System.Runtime.InteropServices.OSPlatform.Windows) Then Return False
                Return ContainsNativeLoadFailure(ex)
            End Function

            Private Shared Function ContainsNativeLoadFailure(ex As Exception) As Boolean
                Dim cur As Exception = ex
                While cur IsNot Nothing
                    If TypeOf cur Is DllNotFoundException Then Return True
                    If TypeOf cur Is TypeInitializationException Then Return True
                    If TypeOf cur Is PlatformNotSupportedException Then Return True
                    cur = cur.InnerException
                End While
                Return False
            End Function

        End Class

    End Class

End Namespace