Option Explicit

Sub Main()

    Dim CAM
    Set CAM = New CAMMaster.Tool

    Dim maxLayers
    maxLayers = 255

    ' ===================== USER INPUT =====================
    Dim folderPath
    folderPath = InputBox("Enter folder path containing ZIP / TGZ files:", _
                          "ZIP Import Folder", "C:\Users\varni\Downloads\Data")

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


    ' ===================== LOAD ZIP FILES =====================
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

        Dim ext
        ext = LCase(fso.GetExtensionName(f.Name))

        If ext = "zip" Or ext = "tgz" Then
            fileCount = fileCount + 1
            zipFiles.Add CStr(fileCount), f.Path & "|" & ext
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

    For Each key In zipFiles.Keys

        Dim parts, filePath, fileType
        parts = Split(zipFiles(key), "|")
        filePath = parts(0)
        fileType = parts(1)

        importIndex = importIndex + 1

        ' Find last used layer before import
        beforeMax = 0
        For j = 1 To maxLayers
            If Trim(CAM.LayerName(j)) <> "" Then
                beforeMax = j
            End If
        Next

        ' ---------------- IMPORT ----------------
        If fileType = "zip" Then

            CAM.ImportZip "Zip=" & filePath

        ElseIf fileType = "tgz" Then

            If beforeMax = 0 Then
                CAM.ImportX filePath, "ODB++", 1, 0, 0
            Else
                CAM.CurrentLayer = beforeMax
                CAM.ImportX filePath, "ODB++", beforeMax, 0, 0
            End If

        End If
        ' ----------------------------------------

        ' Detect newly added layers
        blockStart = 0
        blockEnd = 0

        For j = beforeMax + 1 To maxLayers

            If Trim(CAM.LayerName(j)) <> "" Then

                If blockStart = 0 Then
                    blockStart = j
                End If

                blockEnd = j

            Else

                If blockStart <> 0 Then
                    Exit For
                End If

            End If

        Next

        If blockStart = 0 Then
            GoTo SkipThisZip
        End If

        ' Calculate grid position
        Dim idx0, row, col
        Dim moveX, moveY

        idx0 = importIndex - 1
        row = Int(idx0 / perRow)
        col = idx0 Mod perRow

        moveX = col * spacingX
        moveY = row * spacingY

        ' ===================== MOVE EACH LAYER =====================
        For L = blockStart To blockEnd

            CAM.ClearSelection
            CAM.OnlyCurrentLayer = True
            CAM.OnlyCurrentLayer = False

            ' 🔴 FIX: set current layer BEFORE select
            CAM.CurrentLayer = L

            CAM.SelectAll
            CAM.StepAndRepeatUngroup(0)

            CAM.ClearSelection
            CAM.OnlyCurrentLayer = True

            CAM.ClearSelection
            CAM.OnlyCurrentLayer = True
            CAM.CurrentLayer = L

            CAM.SelectEx "New", "Frame", 0, 0, 6000, 6000
            CAM.MoveSelected moveX, moveY

            CAM.ClearSelection

        Next

        CAM.OnlyCurrentLayer = False

SkipThisZip:

    Next


    ' ===================== COLLECT UNIQUE SUFFIXES =====================
    Dim dictSeen, dictOrder
    Set dictSeen = CreateObject("Scripting.Dictionary")
    Set dictOrder = CreateObject("Scripting.Dictionary")

    Dim lname, dashPos, suffix
    Dim ucount
    ucount = 0

    For j = 1 To maxLayers

        lname = Trim(CAM.LayerName(j))
        If lname <> "" Then

            dashPos = InStrRev(lname, "-")
            If dashPos > 0 Then
                suffix = Mid(lname, dashPos + 1)
            Else
                suffix = lname
            End If

            suffix = LCase(Trim(suffix))

            If Not dictSeen.Exists(suffix) Then
                ucount = ucount + 1
                dictSeen.Add suffix, True
                dictOrder.Add CStr(ucount), suffix
            End If

        End If

    Next

    ' (everything below remains IDENTICAL to your original file)

    ' ===================== MERGE =====================
    CAM.CombineLayersByBoardNumber()

    MsgBox "Completed Successfully!" & vbCrLf & _
           "Files imported: " & fileCount & vbCrLf & _
           "Unique layers: " & ucount

End Sub
