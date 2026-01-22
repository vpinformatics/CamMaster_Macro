Imports System.Runtime.InteropServices
Imports System.IO
Imports CamMasterComTest.License


<ComVisible(True)>
<Guid("B1C1A9F3-0C59-4E4F-9F29-CC4E63F2E111")>
<ClassInterface(ClassInterfaceType.None)>
<ProgId("CamMasterComTest.ZipProcessor")>
Public Class ZipProcessor
    Implements IZipProcessor

    Public Sub RunZipMacro(folderPath As String,
                           perRow As Integer,
                           spacingX As Integer,
                           spacingY As Integer) _
        Implements IZipProcessor.RunZipMacro

        ' 🔒 LICENSE CHECK
        Try
            License.Validate()
        Catch ex As Exception
            MsgBox(ex.Message)
            Exit Sub
        End Try

        Dim CAM As Object
        CAM = CreateObject("CAMMaster.Tool")

        Dim maxLayers As Integer = 255

        ' ===================== LOAD ZIP FILES =====================
        Dim zipFiles As New Dictionary(Of Integer, String)
        Dim fileCount As Integer = 0

        For Each f In Directory.GetFiles(folderPath, "*.zip")
            fileCount += 1
            zipFiles.Add(fileCount, f)
        Next

        If fileCount = 0 Then
            MsgBox("No ZIP files found.")
            Exit Sub
        End If

        ' ===================== IMPORT + GRID MOVE =====================
        Dim importIndex As Integer = 0

        For Each kv In zipFiles

            importIndex += 1

            ' Find last used layer before import
            Dim beforeMax As Integer = 0
            For j = 1 To maxLayers
                If Trim(CAM.LayerName(j)) <> "" Then
                    beforeMax = j
                End If
            Next

            ' Import ZIP
            CAM.ImportZip("Zip=" & kv.Value)

            ' Detect newly added layers
            Dim blockStart As Integer = 0
            Dim blockEnd As Integer = 0

            For j = beforeMax + 1 To maxLayers
                If Trim(CAM.LayerName(j)) <> "" Then
                    If blockStart = 0 Then blockStart = j
                    blockEnd = j
                Else
                    If blockStart <> 0 Then Exit For
                End If
            Next

            If blockStart = 0 Then Continue For

            Dim idx0 As Integer = importIndex - 1
            Dim row As Integer = idx0 \ perRow
            Dim col As Integer = idx0 Mod perRow

            Dim moveX As Integer = col * spacingX
            Dim moveY As Integer = row * spacingY

            ' Move each layer safely (EXACT COPY)
            For L = blockStart To blockEnd

                CAM.ClearSelection()
                CAM.OnlyCurrentLayer = True
                CAM.OnlyCurrentLayer = False
                CAM.SelectAll()
                CAM.StepAndRepeatUngroup(0)
                CAM.ClearSelection()

                CAM.OnlyCurrentLayer = True
                CAM.CurrentLayer = L

                CAM.SelectEx("New", "Frame", -50, -50, 6000, 6000)
                CAM.MoveSelected(moveX, moveY)
                CAM.ClearSelection()

            Next

            CAM.OnlyCurrentLayer = False

        Next

        ' ===================== COLLECT UNIQUE SUFFIXES =====================
        Dim dictSeen As New Dictionary(Of String, Boolean)
        Dim dictOrder As New Dictionary(Of Integer, String)
        Dim ucount As Integer = 0

        For j = 1 To maxLayers
            Dim lname As String = Trim(CAM.LayerName(j))
            If lname <> "" Then
                Dim dashPos = lname.LastIndexOf("-"c)
                Dim suffix As String = If(dashPos >= 0, lname.Substring(dashPos + 1), lname)
                suffix = suffix.ToLower().Trim()
                If Not dictSeen.ContainsKey(suffix) Then
                    ucount += 1
                    dictSeen.Add(suffix, True)
                    dictOrder.Add(ucount, suffix)
                End If
            End If
        Next

        ' ===================== REORDER CANONICAL LAYERS =====================
        For targetIndex = 1 To ucount
            Dim currentSuffix = dictOrder(targetIndex)
            Dim foundIndex As Integer = 0

            For j = 1 To maxLayers
                Dim lname As String = Trim(CAM.LayerName(j))
                If lname <> "" Then
                    Dim dashPos = lname.LastIndexOf("-"c)
                    Dim suffix As String = If(dashPos >= 0, lname.Substring(dashPos + 1), lname)
                    If suffix.ToLower().Trim() = currentSuffix Then
                        foundIndex = j
                        Exit For
                    End If
                End If
            Next

            If foundIndex > 0 AndAlso foundIndex <> targetIndex Then
                CAM.LayerMove(foundIndex, targetIndex)
            End If
        Next

        ' ===================== INSERT ONE BLANK LAYER =====================
        Dim blankPos As Integer = ucount + 1
        If Trim(CAM.LayerName(blankPos)) <> "" Then
            For j = maxLayers To blankPos + 1 Step -1
                If Trim(CAM.LayerName(j)) = "" Then
                    CAM.LayerMove(blankPos, j)
                    Exit For
                End If
            Next
        End If

        ' ===================== MAP B-LAYERS =====================
        Dim mapSuffix As New Dictionary(Of String, Integer)

        For j = 1 To ucount
            Dim lname As String = Trim(CAM.LayerName(j))
            Dim dashPos = lname.LastIndexOf("-"c)
            Dim suffix As String = If(dashPos >= 0, lname.Substring(dashPos + 1), lname)
            suffix = suffix.ToLower().Trim()
            If Not mapSuffix.ContainsKey(suffix) Then
                mapSuffix.Add(suffix, j)
            End If
        Next

        For j = 1 To maxLayers
            Dim lname As String = Trim(CAM.LayerName(j))
            If lname <> "" Then
                Dim dashPos = lname.LastIndexOf("-"c)
                Dim suffix As String = If(dashPos >= 0, lname.Substring(dashPos + 1), lname)
                suffix = suffix.ToLower().Trim()
                If mapSuffix.ContainsKey(suffix) Then
                    CAM.LayerBoardNum(j) = mapSuffix(suffix)
                End If
            End If
        Next

        ' ===================== MERGE =====================
        CAM.CombineLayersByBoardNumber()

        ' ===================== RENAME =====================
        Dim userPrefix As String = InputBox(
    "Enter prefix for layer names:",
    "Rename",
    "ClientX"
)

        If userPrefix IsNot Nothing AndAlso userPrefix.Trim() <> "" Then

            For j As Integer = 1 To maxLayers

                Dim lname As String = CAM.LayerName(j)
                If lname IsNot Nothing AndAlso lname.Trim() <> "" Then

                    Dim dashPos As Integer = lname.LastIndexOf("-"c)
                    Dim suffix As String

                    If dashPos >= 0 Then
                        suffix = lname.Substring(dashPos + 1)
                    Else
                        suffix = lname
                    End If

                    CAM.LayerName(j) = userPrefix & "-" & suffix.Trim()

                End If

            Next

        End If


        ' ===================== CLEANUP =====================
        For j As Integer = ucount + 2 To maxLayers

            Dim lname As String = CAM.LayerName(j)
            If lname IsNot Nothing AndAlso lname.Trim() <> "" Then
                CAM.LayerDelete(j, "All", False)
                CAM.LayerName(j) = ""
            End If

        Next





        MsgBox("ZIP Processing Completed Successfully." & vbCrLf &
               "ZIP files: " & fileCount & vbCrLf &
               "Unique layers: " & ucount)

    End Sub

End Class
