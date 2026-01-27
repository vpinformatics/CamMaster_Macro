Imports System.Runtime.InteropServices
Imports System.IO
'Imports System.Text.Json
Imports CamMasterComTest.License
Imports Newtonsoft.Json


<ComVisible(True)>
<Guid("B1C1A9F3-0C59-4E4F-9F29-CC4E63F2E111")>
<ClassInterface(ClassInterfaceType.None)>
<ProgId("CamMasterComTest.ZipProcessor")>
Public Class ZipProcessor
    Implements IZipProcessor

    Private Class PlacementItem
        Public Property x As Double
        Public Property y As Double
        Public Property label As String
        Public Property file As String
        Public Property rotation As Double
    End Class

    Private Class JsonRoot
        Public Property objects As List(Of PlacementItem)
    End Class

    Public Sub RunZipMacro(jsonFilePath As String) _
        Implements IZipProcessor.RunZipMacro

        ' ===================== LICENSE CHECK =====================
        Try
            License.Validate()
        Catch ex As Exception
            MsgBox(ex.Message)
            Exit Sub
        End Try

        If Not File.Exists(jsonFilePath) Then
            MsgBox("JSON file not found.")
            Exit Sub
        End If

        Dim jsonText As String = File.ReadAllText(jsonFilePath)
        'Dim jsonData As JsonRoot =
        '    JsonSerializer.Deserialize(Of JsonRoot)(jsonText)

        Dim jsonData As JsonRoot =
        JsonConvert.DeserializeObject(Of JsonRoot)(jsonText)

        If jsonData Is Nothing OrElse jsonData.objects Is Nothing OrElse jsonData.objects.Count = 0 Then
            MsgBox("No placement objects found in JSON.")
            Exit Sub
        End If

        Dim baseFolder As String = Path.GetDirectoryName(jsonFilePath)

        Dim CAM As Object = CreateObject("CAMMaster.Tool")
        CAM.Units = "mm"
        Dim maxLayers As Integer = 255
        Dim fileCount As Integer = 0

        ' ===================== IMPORT + PLACE FROM JSON =====================
        For Each item In jsonData.objects

            Dim zipPath As String = Path.Combine(baseFolder, item.file)

            If Not File.Exists(zipPath) Then
                MsgBox("ZIP file not found: " & zipPath)
                Continue For
            End If

            fileCount += 1

            ' Find last used layer BEFORE import
            Dim beforeMax As Integer = 0
            For j = 1 To maxLayers
                If Trim(CAM.LayerName(j)) <> "" Then
                    beforeMax = j
                End If
            Next

            ' Import ZIP
            CAM.ImportZip("Zip=" & zipPath)

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

            Dim bs = blockStart

            ' ===================== ROTATE + PLACE (ZIP BLOCK AS ONE UNIT) =====================

            ' 1) Build layer list for this ZIP
            Dim layerList As New List(Of String)
            For L = blockStart To blockEnd
                layerList.Add(L.ToString())
            Next
            Dim layerCsv As String = String.Join(",", layerList)

            ' 2) Select all layers of this ZIP together
            CAM.ClearSelection()
            CAM.SelectLayers(layerCsv)

            ' 3) Select all geometry from selected layers
            CAM.SelectEx("New", "Frame", -7500, -7500, 7500, 7500)

            ' 4) Get bottom-left
            Dim bl1 = GetBottomLeftFromSelection(CAM)

            ' 5) Move bottom-left to cursor (0,0)
            CAM.MoveSelected(-bl1.Item1, -bl1.Item2)

            ' 6) Rotate ONCE around cursor
            If item.rotation <> 0 Then
                CAM.RotateSelectedEx(item.rotation, 1, "")
            End If

            ' 7) Recalculate bottom-left after rotation
            Dim bl2 = GetBottomLeftFromSelection(CAM)

            ' 8) Move entire ZIP block to target (x,y)
            CAM.MoveSelected(item.x - bl2.Item1, item.y - bl2.Item2)

            CAM.ClearSelection()


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
                    Dim suffix As String = If(dashPos >= 0, lname.Substring(dashPos + 1), lname)
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

        MsgBox(
            "ZIP Processing Completed Successfully." & vbCrLf &
            "ZIP files: " & fileCount & vbCrLf &
            "Unique layers: " & ucount
        )

    End Sub

    Private Function GetBottomLeftFromSelection(CAM As Object) As Tuple(Of Double, Double)

        Dim centersText As String = CAM.GetSelectedElementsCenters()
        Dim lines() As String = centersText.Split({vbCrLf, vbLf}, StringSplitOptions.RemoveEmptyEntries)

        Dim minX As Double = Double.MaxValue
        Dim minY As Double = Double.MaxValue

        For Each ln As String In lines
            Dim parts() As String = ln.Trim().Split({" "c, vbTab}, StringSplitOptions.RemoveEmptyEntries)
            If parts.Length >= 3 Then
                Dim cx As Double = CDbl(parts(1))
                Dim cy As Double = CDbl(parts(2))
                If cx < minX Then minX = cx
                If cy < minY Then minY = cy
            End If
        Next

        If minX = Double.MaxValue Then
            Return Tuple.Create(0.0, 0.0)
        End If

        Return Tuple.Create(minX, minY)

    End Function


End Class




