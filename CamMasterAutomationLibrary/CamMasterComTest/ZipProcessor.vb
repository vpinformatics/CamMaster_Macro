Imports System.Runtime.InteropServices
Imports System.IO
'Imports System.Text.Json
Imports CamMasterComTest.License
Imports Newtonsoft.Json
Imports System.Windows.Forms
Imports System.Threading

<ComVisible(True)>
Public Class KeySender

    Public Sub Send(keys As String)
        Dim t As New Thread(
            Sub()
                SendKeys.SendWait(keys)
            End Sub)

        t.SetApartmentState(ApartmentState.STA)
        t.Start()
        t.Join()
    End Sub

End Class



<ComVisible(True)>
<Guid("B1C1A9F3-0C59-4E4F-9F29-CC4E63F2E111")>
<ClassInterface(ClassInterfaceType.None)>
<ProgId("CamMasterComTest.ZipProcessor")>
Public Class ZipProcessor
    Implements IZipProcessor

    Dim DCodeNumber = 10000

    Private Class PlacementItem
        Public Property x As Double
        Public Property y As Double
        Public Property label As String
        Public Property file As String
        Public Property rotation As Double
        Public Property hasOuterBorder As Boolean


        Public Property camX As Double
        Public Property camY As Double

        Public Property comboID As String
        Public Property launchingPanels As Integer


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
        Dim counter As Integer = 1
        Dim uniqueLayers As Integer = 0

        'Dim jsonFile As String = jsonFilePath _
        '        .Replace(baseFolder, "") _
        '        .Replace("/", "") _
        '        .Replace("\", "") _
        '        .Replace(".json", "") _
        '        .ToUpper()

        'Dim comboID = jsonFile


        Dim fileNameOnly As String = Path.GetFileNameWithoutExtension(jsonFilePath)

        Dim parts() As String = fileNameOnly.Split(New Char() {"_"c, "-"c}, StringSplitOptions.RemoveEmptyEntries)

        Dim comboID As String = ""
        Dim launchingPanels As Integer = 1   ' ✅ DEFAULT = 1

        If parts.Length >= 1 Then
            comboID = parts(0).ToUpper().Trim()
        End If

        If parts.Length >= 2 Then
            Integer.TryParse(parts(1), launchingPanels)
        End If
        Dim jsonFile = comboID


        ' ===================== RENAME =====================
        Dim userPrefix As String = InputBox(
            "Enter prefix for layer names:",
            "Rename",
            jsonFile)


        Dim a = 1

        If a = 1 Then

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
                'If uniqueLayers > 0 Then
                '    beforeMax = uniqueLayers
                'End If
                For j = 1 To maxLayers
                    If Trim(CAM.LayerName(j)) <> "" Then beforeMax = j
                Next


                ' Import ZIP
                CAM.CurrentLayer = uniqueLayers + 12
                CAM.ImportZip("Zip=" & zipPath)
                ' AppActivate("Scan Results")
                'Thread.Sleep(1500)

                'Dim x = New KeySender()
                'x.Send("{ENTER}")
                'System.Windows.Forms.SendKeys.SendWait("{ENTER}")


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


                'Dim diff As Integer = blockEnd - blockStart
                'If (uniqueLayers > 0) Then
                '    blockStart = uniqueLayers + 1
                '    blockEnd = blockStart + diff
                'End If
                'MsgBox(blockStart.ToString() & "-" & blockEnd.ToString())





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
                CAM.OnlyCurrentLayer = False
                CAM.SelectEx("New", "Frame", -100000, -100000, 100000, 100000)
                CAM.StepAndRepeatUngroup(0)

                ' 4) Get bottom-left
                Dim bl1 = GetBottomLeftFromSelection(CAM)

                ' 5) Move bottom-left to cursor (0,0)
                CAM.MoveSelected(-bl1.Item1, -bl1.Item2)

                ' 6) Rotate ONCE around cursor
                If item.rotation <> 0 Then
                    CAM.RotateSelectedEx(item.rotation, 0, "")
                End If

                ' 7) Recalculate bottom-left after rotation
                Dim bl2 = GetBottomLeftFromSelection(CAM)

                ' 8) Move entire ZIP block to target (x,y)

                '============================Old Version Settings=======================
                Dim penMargin As Decimal = 0 '0.3048
                Dim padding As Decimal = 12
                Dim outerBorder As Decimal = 1
                If item.hasOuterBorder = True Then
                    CAM.MoveSelected(item.x - bl2.Item1 - outerBorder - padding - penMargin, item.y - bl2.Item2 - outerBorder - padding - penMargin)
                Else
                    CAM.MoveSelected(item.x - bl2.Item1 - padding - penMargin, item.y - bl2.Item2 - padding - penMargin)
                End If

                '============================New Version Settings=======================
                'Dim penMargin As Decimal = 0 '0.3048
                'Dim padding As Decimal = 12
                'Dim outerBorder As Decimal = 0
                'CAM.MoveSelected(item.camX - bl2.Item1 - padding - penMargin, item.camY - bl2.Item2 - padding - penMargin)
                'CAM.ClearSelection()







                If counter > 1 Then
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

                    If userPrefix IsNot Nothing AndAlso userPrefix.Trim() <> "" Then
                        For j = 1 To maxLayers
                            Dim lname As String = CAM.LayerName(j)
                            If lname IsNot Nothing AndAlso lname.Trim() <> "" Then
                                Dim dashPos = lname.LastIndexOf("-"c)
                                Dim suffix As String = If(dashPos >= 0, lname.Substring(dashPos + 1), lname)
                                CAM.LayerName(j) = userPrefix & "-" & suffix.Trim()
                            End If
                        Next
                    End If

                    ' ===================== CLEANUP =====================
                    For j = ucount + 2 To maxLayers
                        Dim lname As String = CAM.LayerName(j)
                        If lname IsNot Nothing AndAlso lname.Trim() <> "" Then
                            CAM.LayerDelete(j, "All", False)
                            CAM.LayerName(j) = ""
                        End If
                    Next
                    uniqueLayers = ucount





                End If
                counter = counter + 1

            Next
        End If



        ' ===================== Adding FRAMES=====================
        If a = 1 Then


            ' ===================== Adding FRAMES=====================
            Dim zipPathFrame As String = Path.Combine(baseFolder, "Frame.zip")

            If Not File.Exists(zipPathFrame) Then
                MsgBox("ZIP file not found: " & zipPathFrame)
                GoTo EndFrame
            End If

            ' Find last used layer BEFORE import
            Dim beforeMaxFrame As Integer = 0
            For j = 1 To maxLayers
                If Trim(CAM.LayerName(j)) <> "" Then beforeMaxFrame = j
            Next


            ' Import ZIP
            CAM.CurrentLayer = uniqueLayers + 12
            CAM.ImportZip("Zip=" & zipPathFrame)

            ' Detect newly added layers
            Dim blockStartFrame As Integer = 0
            Dim blockEndFrame As Integer = 0

            For j = beforeMaxFrame + 1 To maxLayers
                If Trim(CAM.LayerName(j)) <> "" Then
                    If blockStartFrame = 0 Then blockStartFrame = j
                    blockEndFrame = j
                Else
                    If blockStartFrame <> 0 Then Exit For
                End If
            Next

            If blockStartFrame = 0 Then GoTo EndFrame


            ' ===================== ROTATE + PLACE (ZIP BLOCK AS ONE UNIT) =====================

            ' 1) Build layer list for this ZIP
            Dim layerListFrame As New List(Of String)
            For L = blockStartFrame To blockEndFrame
                layerListFrame.Add(L.ToString())
            Next
            Dim layerCsvFrane As String = String.Join(",", layerListFrame)

            ' 2) Select all layers of this ZIP together
            CAM.ClearSelection()
            CAM.SelectLayers(layerCsvFrane)

            ' 3) Select all geometry from selected layers
            CAM.OnlyCurrentLayer = False
            CAM.SelectEx("New", "Frame", -100000, -100000, 100000, 100000)
            CAM.StepAndRepeatUngroup(0)


            CAM.ClearSelection()







            If counter > 1 Then
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

                '======================Text on frames========================
                ' D CODE SETTING
                'CAM.SetupDCodesDialog()
                CAM.DCodeShapeAndAngle(DCodeNumber) = "C, Diameter=0.2, Hole=0"

                Dim todayDate As String = Date.Now.ToString("dd-MM-yyyy")
                For j = 1 To maxLayers

                    'For j = 1 To 20

                    Dim lname As String = Trim(CAM.LayerName(j))
                    'If lname = "" Then Continue For

                    ' MsgBox(lname.ToUpper())


                    CAM.AllLayersVis("On")
                    CAM.AllLayersVis("Toggle")
                    CAM.CurrentLayer = j
                    CAM.LayerVisible(j) = True


                    Select Case lname.ToUpper()

                        Case "FRAME-TOPCKT.274X"

                            InsertLayerTextByIndex(
                    CAM, j,
                    "CIRCUITWALA " & comboID & " TOP CKT DATE : " & todayDate & " QTY : " & launchingPanels,
                    False
                )

                        Case "FRAME-TOPMASK.274X"
                            InsertLayerTextByIndex(
                    CAM, j,
                    "CIRCUITWALA " & comboID & " TOP MASK DATE : " & todayDate & " QTY : " & launchingPanels,
                    False
                )

                        Case "FRAME-TOPSILK.274X"
                            InsertLayerTextByIndex(
                    CAM, j,
                    "CIRCUITWALA " & comboID & " TOP SILK DATE : " & todayDate & " QTY : " & launchingPanels,
                    False
                )

                        Case "FRAME-BOTCKT.274X"
                            InsertLayerTextByIndex(
                    CAM, j,
                    "CIRCUITWALA " & comboID & " BOT CKT DATE : " & todayDate & " QTY : " & launchingPanels,
                    True
                )

                        Case "FRAME-BOTMASK.274X"
                            InsertLayerTextByIndex(
                    CAM, j,
                    "CIRCUITWALA " & comboID & " BOT MASK DATE : " & todayDate & " QTY : " & launchingPanels,
                    True
                )

                        Case "FRAME-BOTSILK.274X"
                            InsertLayerTextByIndex(
                    CAM, j,
                    "CIRCUITWALA " & comboID & " BOT SILK DATE : " & todayDate & " QTY : " & launchingPanels,
                    True
                )

                    End Select

                    CAM.AllLayersVis("On")

                Next


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

                If userPrefix IsNot Nothing AndAlso userPrefix.Trim() <> "" Then
                    For j = 1 To maxLayers
                        Dim lname As String = CAM.LayerName(j)
                        If lname IsNot Nothing AndAlso lname.Trim() <> "" Then
                            Dim dashPos = lname.LastIndexOf("-"c)
                            Dim suffix As String = If(dashPos >= 0, lname.Substring(dashPos + 1), lname)
                            CAM.LayerName(j) = userPrefix & "-" & suffix.Trim()
                        End If
                    Next
                End If



                ' ===================== CLEANUP =====================
                For j = ucount + 2 To maxLayers
                    Dim lname As String = CAM.LayerName(j)
                    If lname IsNot Nothing AndAlso lname.Trim() <> "" Then
                        CAM.LayerDelete(j, "All", False)
                        CAM.LayerName(j) = ""
                    End If
                Next
                uniqueLayers = ucount


            End If

EndFrame:
        End If

        'CAM.SetupDCodesDialog()
        'CAM.SetupDCodesDialog()
        CAM.DeleteUnusedDCodes()
        CAM.CurrentLayer = 1

        'Next
        '===================================================================
        '===================================================================
        '===================================================================
        '===================================================================
        MsgBox("ZIP Processing Completed Successfully." & vbCrLf &
                    "ZIP files: " & fileCount & vbCrLf &
                    "Unique layers: " & uniqueLayers)



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

    Dim todayDate As String = Date.Now.ToString("dd-MM-yyyy")

    Function BuildFrameText(comboID As String, layerLabel As String, qty As Integer) As String
        Return "CIRCUITWALA " & comboID & " " & layerLabel &
           " DATE : " & todayDate &
           " QTY : " & qty.ToString()
    End Function
    Sub InsertLayerText(CAM As CAMMaster.Tool, layerFileName As String, frameText As String, isMirrored As Boolean)

        CAM.SetCurrentLayerByName(layerFileName)

        ' TEXT PARAMETERS (AS REQUIRED)
        CAM.SetTextParameters(
        "Height=2.5," &
        "Font=Uppercase," &
        "AspectRatio=1," &
        "Angle=0," &
        "Mirrored=" & isMirrored.ToString() & "," &
        "CharSpacing=0.5," &
        "LineSpacing=0"
    )

        ' FRAME POSITION
        CAM.MoveCursor(-60, 160)

        ' INSERT TEXT
        CAM.InsertText(frameText, DCodeNumber, True)

        ' ENSURE MIRROR FOR BOTTOM
        If isMirrored Then
            CAM.MirrorSelectedEx(False, False, "")
        End If

        CAM.ClearSelection()

    End Sub
    Sub InsertLayerTextByIndex(CAM As CAMMaster.Tool, layerIndex As Integer, frameText As String, mirrored As Boolean)

        CAM.CurrentLayer = layerIndex

        'CAM.SetupDCodesDialog()
        'CAM.DCodeShapeAndAngle(DCodeNumber) = "C, Diameter=0.25, Hole=0"

        CAM.SetTextParameters(
        "Height=2," &
        "Font=Uppercase," &
        "AspectRatio=1," &
        "Angle=0," &
        "Mirrored=" & mirrored.ToString() & "," &
        "CharSpacing=0.5," &
        "LineSpacing=0"
    )

        CAM.MoveCursor(-8.5, 107)
        CAM.InsertText(frameText, DCodeNumber, True)

        ' Add text at origin
        ' CAM.TextAdd(0, 0, panelText)

        ' Select new text near origin
        CAM.ClearSelection()
        'CAM.SelectEx("New", "Frame", -13, 100, 115, 115)
        CAM.SelectEx("New", "Frame", -100, 100, 115, 115)

        ' Change font
        ' CAM.TextFont("Arial")

        ' Rotate text
        CAM.RotateSelectedEx(90, 1, "")

        ' Move to desired coordinate
        CAM.MoveSelected(2, 2)

        If mirrored Then
            'CAM.MirrorSelectedEx(False, False, "")
            CAM.MoveSelected(0, 88)
        End If

        CAM.ClearSelection()



        If mirrored Then
            CAM.MirrorSelectedEx(False, False, "")
        End If

        CAM.ClearSelection()



    End Sub
End Class
