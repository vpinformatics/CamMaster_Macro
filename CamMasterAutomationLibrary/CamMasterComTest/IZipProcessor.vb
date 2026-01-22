Imports System.Runtime.InteropServices

<ComVisible(True)>
<Guid("8C6F7F4E-9F61-4D2F-8A72-91C62B8E3D10")>
<InterfaceType(ComInterfaceType.InterfaceIsIDispatch)>
Public Interface IZipProcessor
    Sub RunZipMacro(folderPath As String, perRow As Integer, spacingX As Integer, spacingY As Integer)
End Interface
