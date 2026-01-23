Option Explicit

Sub Main()

    Dim jsonPath
    jsonPath = InputBox( _
        "Enter JSON file path:", _
        "JSON Placement File", _
        "C:\Users\varni\Downloads\Data\layout.json" _
    )

    If jsonPath = "" Then Exit Sub

    Dim obj
    Set obj = CreateObject("CamMasterComTest.ZipProcessor")

    obj.RunZipMacro jsonPath

End Sub
