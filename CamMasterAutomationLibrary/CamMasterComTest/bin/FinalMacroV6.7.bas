Option Explicit

Sub Main()

    'Dim Shell, folder, folderPath
    'Set Shell = CreateObject("Shell.Application")
'
    'Set folder = Shell.BrowseForFolder(0, _
        '"Select folder containing ZIP files", _
        '0, _
        '17)   ' This PC
'
    'If folder Is Nothing Then Exit Sub
    'folderPath = folder.Self.Path

    ' ===================== USER INPUT =====================
    Dim folderPath
    folderPath = InputBox("Enter folder path containing ZIP files:", _
                          "ZIP Import Folder", "C:\Users\varni\Downloads\Data")


    Dim obj
    Set obj = CreateObject("CamMasterComTest.ZipProcessor")

    ' SAME VALUES AS YOUR MACRO
    obj.RunZipMacro folderPath, 4, 300, 300


End Sub
