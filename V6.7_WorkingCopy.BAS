Option Explicit

Sub Main()

    Dim CAM
    Set CAM = New CAMMaster.Tool

    Dim maxLayers
    maxLayers = 255

    ' ===================== USER INPUT =====================
    Dim folderPath
    folderPath = InputBox("Enter folder path containing ZIP / TGZ files:", _
                          "Import Folder", "C:\Users\varni\Downloads\Data")

    If Trim(folderPath) = "" Then
        MsgBox "No folder selected. Exiting."
        Exit Sub
    End If

    Dim perRow
    perRow = 4

    Dim spacingX, spacingY
    spacingX = 300
    spacingY = 300
    ' =====================================================


    ' ===================== LOAD ZIP / TGZ FILES =====================
    Dim fso, fld, f
    Set fso = CreateObject("Scripting.FileSystemObject")

    On Error Resume Next
    Set fld = fso.GetFolder(folderPath)
    If Err.Number <> 0 Then
        MsgBox "Folder not found."
        Exit Sub
    End If
    On Error GoTo 0

    Dim zipFiles
    Set zipFiles = CreateObject("Scripting.Dictionary")

    Dim fileCount
    fileCount = 0

    For Each f In fld.Files
        If LCase(fso.GetExtensionName(f.Name)) = "zip" _
        Or LCase(fso.GetExtensionName(f.Name)) = "tgz" Then
            fileCount = fileCount + 1
            zipFiles.Add CStr(fileCount), f.Path
        End If
    Next

    If fileCount = 0 Then
        MsgBox "No ZIP or TGZ files found."
        Exit Sub
    End If


    ' ===================== IMPORT + GRID MOVE =====================
    Dim importIndex
    importIndex = 0

    Dim beforeMax, blockStart, blockEnd
    Dim j, L, key
    Dim lname

    For Each key In zipFiles.Keys

        importIndex = importIndex + 1

        beforeMax = 0
        For j = 1 To maxLayers
            If Trim(CAM.LayerName(j)) <> "" Then beforeMax = j
        Next

        Dim ext
        ext = LCase(fso.GetExtensionName(zipFiles(key)))

        ' ===================== IMPORT =====================
        If ext = "zip" Then
            CAM.ImportZip "Zip=" & zipFiles(key)
        ElseIf ext = "tgz" Then
            CAM.ImportX zipFiles(key), "ODB++", 1, 0, 0
        End If


        ' ===================== DETECT NEW LAYERS =====================
        blockStart = 0
        blockEnd = 0

        For j = beforeMax + 1 To maxLayers
            If Trim(CAM.LayerName(j)) <> "" Then
                If blockStart = 0 Then blockStart = j
                blockEnd = j
            Else
                If blockStart <> 0 Then Exit For
            End If
        Next

        If blockStart = 0 Then GoTo SkipThisFile


        ' ===================== GRID POSITION =====================
        Dim idx0, row, col
        Dim moveX, moveY

        idx0 = importIndex - 1
        row = Int(idx0 / perRow)
        col = idx0 Mod perRow

        moveX = col * spacingX
        moveY = row * spacingY


        ' ===================== ZIP MOVE (UNCHANGED) =====================
        If ext = "zip" Then

            For L = blockStart To blockEnd

                CAM.ClearSelection
                CAM.OnlyCurrentLayer = True
                CAM.OnlyCurrentLayer = False
                CAM.SelectAll
                CAM.StepAndRepeatUngroup(0)
                CAM.ClearSelection
                CAM.OnlyCurrentLayer = True

                CAM.ClearSelection
                CAM.CurrentLayer = L
                CAM.SelectEx "New", "Frame", 0, 0, 6000, 6000
                CAM.MoveSelected moveX, moveY
                CAM.ClearSelection

            Next

            CAM.OnlyCurrentLayer = False

        End If


        ' ===================== TGZ MOVE (BOUNDING BOX BASED) =====================
        If ext = "tgz" Then

            Dim outlineLayer
            outlineLayer = 0

            For j = blockStart To blockEnd
                lname = LCase(Trim(CAM.LayerName(j)))
                If InStr(lname, "profile") > 0 _
                Or InStr(lname, "outline") > 0 _
                Or InStr(lname, "bol") > 0 Then
                    outlineLayer = j
                    Exit For
                End If
            Next

            If outlineLayer > 0 Then

                CAM.ClearSelection
                CAM.OnlyCurrentLayer = True
                CAM.CurrentLayer = outlineLayer
                CAM.SelectAll

                Dim dims, a
                dims = CAM.GetSelectionDimensions(False, True)
                CAM.ClearSelection

                If dims <> "" Then
                    a = Split(dims, ",")

                    CAM.SelectEx "New", "Frame", _
                        CDbl(a(0)) - 10, _
                        CDbl(a(1)) - 10, _
                        CDbl(a(2)) + 10, _
                        CDbl(a(3)) + 10

                    CAM.OnlyCurrentLayer = False
                    CAM.MoveSelected moveX, moveY
                    CAM.ClearSelection
                End If

            End If

        End If

SkipThisFile:
    Next


    ' ===================== (REST OF YOUR CODE UNCHANGED) =====================
    ' suffix collection, reorder, merge, rename, cleanup
    ' EXACTLY as in your working version

    MsgBox "V6.6 Completed Successfully!" & vbCrLf & _
           "Files: " & fileCount

End Sub
