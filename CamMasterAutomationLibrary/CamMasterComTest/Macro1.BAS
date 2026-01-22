Option Explicit

Sub Main()

    On Error GoTo ErrHandler

    Dim obj
    Set obj = CreateObject("CamMasterComTest.TestRunner")

    obj.Run

    Set obj = Nothing
    Exit Sub

ErrHandler:
    MsgBox "DLL call failed:" & vbCrLf & Err.Description, vbCritical

End Sub
