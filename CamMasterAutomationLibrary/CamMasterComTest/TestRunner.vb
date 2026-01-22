Imports System.Runtime.InteropServices

<ComVisible(True)>
<Guid("C2E6B5C9-8B5F-4E42-8F5E-4A8F1F7A2B77")>
<ClassInterface(ClassInterfaceType.None)>
<ProgId("CamMasterComTest.TestRunner")>
Public Class TestRunner
    Implements ITestRunner

    Public Sub Run() Implements ITestRunner.Run
        MsgBox("Basic COM call OK")
    End Sub

    Public Sub RunWithParams(folderPath As String, perRow As Integer, spacingX As Integer, spacingY As Integer) _
    Implements ITestRunner.RunWithParams

        License.Validate()   ' 🔒 HARD STOP if machine not allowed

        MsgBox("License OK. Params accepted.", vbInformation)

    End Sub


End Class
