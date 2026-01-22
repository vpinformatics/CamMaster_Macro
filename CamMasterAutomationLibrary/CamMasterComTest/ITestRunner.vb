Imports System.Runtime.InteropServices

<ComVisible(True)>
<Guid("6B9A0B7F-6C9E-4A0A-9F65-90D5A0E6C112")>
<InterfaceType(ComInterfaceType.InterfaceIsIDispatch)>
Public Interface ITestRunner

    Sub Run()
    Sub RunWithParams(folderPath As String, perRow As Integer, spacingX As Integer, spacingY As Integer)

End Interface
