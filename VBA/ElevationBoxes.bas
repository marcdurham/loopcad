Attribute VB_Name = "ElevationBoxes"
Public Function GetElevationBoxFloorTag(testPoint As Variant) As FloorTag
Dim eBox As ElevationBox
Dim tag As FloorTag
    
    If FloorTags Is Nothing Then
        ScanEntities
    If FloorTags.Count <= 0 Then
        ScanEntities
    If ElevationBoxes Is Nothing Then
        ScanElevationBoxes
    If ElevationBoxes.Count <= 0 Then
        ScanElevationBoxes
    
    Set eBox = GetElevationBox(testPoint)
    
    If eBox Is Nothing Then
        Exit Function
    
    For Each tag In FloorTags
        If InBox(eBox.Box, tag.insertionPoint) Then
            Set GetElevationBoxFloorTag = tag
            Exit Function
        End If
    Next xFloorTag

End Function
Public Sub ScanFor()
    Dim entit As IntelliCAD.entity
    Dim Boxes As New Collection
    Dim Labels As New Collection
    Dim Box As IntelliCAD.lwPolyline
    Dim Box2d As IntelliCAD.polyline
    Dim eBox As ElevationBox
    Dim innerBox As ElevationBox
    Dim Coordinates As Variant
    Dim ElevLabel As IntelliCAD.text
    Dim ElevMLable As IntelliCAD.mText
    Dim innerBoxCorner(0 To 1) As Double
    Dim BoxUpperLeft(0 To 1) As Double

    'See which Boxes exist
    For Each entit In ActiveDocument.ModelSpace
        'Debug.Print "Elevation Box Layer = "; Entit.Layer
        If UCase(entit.layer) = UCase("ElevationBox") Then
            'Debug.Print "Elevation Box Layer! Entity.EntityName = " & Entit.EntityName
            If UCase(entit.EntityName) = UCase("Polyline") Then
                Set Box = entit
                Coordinates = Box.Coordinates
                If (Coordinates.Count - 1) = 7 Then ' an upper bound of 7 means it's a rectangle
                    'Debug.Print "Box!"
                    Set eBox = New ElevationBox
                    Set eBox.Box = Box
                    Boxes.Add eBox
                End If
            ElseIf UCase(entit.EntityName) = UCase("MText") Then
                Set ElevMLable = entit
                Labels.Add ElevMLable
                
            End If
        End If
    Next entit
    
    'get inner boxes
    For Each eBox In Boxes
        'maybe: If EBox.InnerBoxes Is Nothing Then Set EBox.InnerBoxes = New Collection
        For Each innerBox In Boxes
            innerBoxCorner(0) = innerBox.Box.Coordinates(0)
            innerBoxCorner(1) = innerBox.Box.Coordinates(1)
            If InBox(eBox.Box, innerBoxCorner) _
                And Not (eBox.Box.Coordinates(0) = innerBoxCorner(0) _
                And eBox.Box.Coordinates(1) = innerBoxCorner(1)) Then
                If eBox.InnerBoxes Is Nothing Then Set eBox.InnerBoxes = New Collection
                eBox.InnerBoxes.Add innerBox
            End If
        Next innerBox
    Next eBox
    
    
    If Labels.Count = 0 Then
        MsgBox "The current Elevation Box does not a have a proper Elevation Label in it."
        Err.Raise 1, "PointElevation", "Elevation Box does not have a proper elevation label in it"
        Exit Sub
    End If
        
    'get elevation text
    For Each eBox In Boxes

        For Each ElevMLable In Labels
            If BoxCornerMatch(eBox.Box.Coordinates, ElevMLable.insertionPoint) Then
            'If InBox(EBox.Box, ElevMLable.InsertionPoint) Then
                'Debug.Print "Elevation label is: " & ElevMLable.TextString
                If InStr(1, ElevMLable.TextString, "Elevation ") > 0 Then
                    Dim elev As Double
                    elev = val(Mid(ElevMLable.TextString, InStr(1, ElevMLable.TextString, "Elevation ") + 10))
                    'Debug.Print "elev = " & elev
                    eBox.elevation = elev
                Else
                    eBox.elevation = Null
                End If
            End If
        Next ElevMLable
    Next eBox
    
    Set ElevationBoxes = Boxes
    
End Sub
Public Function GetFrom(testPoint As Variant) As ElevationBox
Dim eBox As ElevationBox
Dim xFloorTag As FloorTag
    
    If ElevationBoxes Is Nothing Then ScanElevationBoxes
    If ElevationBoxes.Count <= 0 Then ScanElevationBoxes
    
    For Each eBox In ElevationBoxes
        If InBox(eBox.Box, testPoint) Then
            Set GetElevationBox = eBox
            Exit Function
        End If
    Next eBox

End Function

Public Sub Draw()
Debug.Print "Preparing to draw Elevation Box..."
On Error GoTo ErrTrap
    Dim startPoint As IntelliCAD.point
    Dim endPoint As IntelliCAD.point
    'Dim newBox(0 To 7) As Double
    Dim newBox As New IntelliCAD.points
    Dim ElevationBox As IntelliCAD.lwPolyline
    Dim textElevation As IntelliCAD.text
    Dim mTextElevation As IntelliCAD.mText
    Dim TextPoint As New IntelliCAD.point
    Dim strElevation As String
    Dim elevation As Double
    Dim upperLeft As New IntelliCAD.point
        
    If LastElevation = 0 Then
        LastElevation = 108
    Else
        LastElevation = LastElevation + 8
    End If
    
    Set startPoint = ActiveDocument.Utility.GetPoint(, "Click the first corner for the Elevation Box")
    Set endPoint = ActiveDocument.Utility.GetCorner(startPoint, "Click the other corner for the Elevation Box")
    
    newBox.Add startPoint.x, startPoint.y
    newBox.Add endPoint.x, startPoint.y
    newBox.Add endPoint.x, endPoint.y
    newBox.Add startPoint.x, endPoint.y
    
    strElevation = InputBox("Input elevation in feet", "DrawElevationBox", LastElevation)
    If strElevation <> "" Then
        'CheckAddLayer "ElevationBox", , False, True
        
        elevation = val(strElevation)
    
        Set ElevationBox = ActiveDocument.ModelSpace.AddLightWeightPolyline(newBox)
        
        ElevationBox.layer = "ElevationBox"
        ElevationBox.LineType = "ByLayer"
        ElevationBox.Closed = True
        ElevationBox.Color.ColorIndex = Colors.vicMagenta
        
        'TextPoint.x =BoxUpperLeft(newBox).x
        'TextPoint.y = BoxUpperLeft(newBox).y
        
        Set mTextElevation = ActiveDocument.ModelSpace.AddMText(startPoint, 108, "Elevation " & elevation)
        mTextElevation.layer = "ElevationBox"
        mTextElevation.Color.ColorIndex = Colors.vicMagenta
    Else
        MsgBox "Invalid elevation!"
    End If
    
ExitProc:
    Exit Sub
ErrTrap:
    If Err.Number = -2147352567 Then
        Resume ExitProc
    Else
        MsgBox Err.Number & " " & Err.Description, vbCritical, "Scanner.DrawElevationBox()"
        Resume ExitProc
        Resume
    End If

End Sub

