Attribute VB_Name = "Scanner"
Option Explicit
'Dim Things As New Collection
Public Job As Job
Public nodes As New Collection

Public ElevationBoxes As Collection

Public Heads As New Collection
Public Risers As New Collection
Public Tees As New Collection
Public Fixtures As New Collection
Public ManifoldNodes As New Collection

Public Pipes As New Collection
Public NodeIDs As New Collection  'new 2010-08-23
Public HeadPairs As New Collection 'new...

Public FloorTags As New Collection
Public FloorConnectors As New Collection

Public headModels As Collection
'Public conLoopCalcData As ADODB.Connection

Public Const ICON_FOLDER = "D:\LoopCAD\Icons"
Public Const LISP_FOLDER = "D:\LoopCAD"
Public Const SUPPORT_FOLDER = "D:\LoopCAD"
Public Const NODE_TYPE_HEAD = 2
Public Const NODE_TYPE_TEE = 3
Public Const NODE_TYPE_RISER = 4
Public Const NODE_TYPE_MANIFOLD_PORT = 5
Public Const NODE_TYPE_FIXTURE = 6
Public Const PI = 3.14159265
Public Const LINE_POINT_TOLERENCE = 0.0000001
Public Const POINTS_MATCH_ALMOST = 0.00001
Public Const COLOR_CYAN = IntelliCAD.Colors.vicCyan

Public Enum PipeColors
    Red = 1
    Cyan = 2
    Magenta = 3
    Green = 4
    Blue = 5
End Enum

Public Type Vertex
    x As Double
    y As Double
    Z As Double
End Type
Public LastElevation As Long
Public ErrorCircleCount As Long













Public Sub AlignRisers()
Dim xRiser As Riser
Dim yRiser As Riser
Dim closestFloorTag As FloorTag
Dim relative As Variant
Dim xFloorPoint As Variant '(0 To 2) As Double
Dim yFloorPoint As Variant '(0 To 2) As Double

    If Risers.Count > 0 Then
        For Each xRiser In Risers
            '''If xRiser.Name Like "*.L" Then '**** test only
                xFloorPoint = FloorCoordinates(xRiser.insertionPoint)
                Debug.Print xRiser.Name & ": Relative " & xFloorPoint(0) & ", " & xFloorPoint(1) & " Absolute " & xRiser.insertionPoint.x & ", " & xRiser.insertionPoint.y
                For Each yRiser In Risers
                    yFloorPoint = FloorCoordinates(yRiser.insertionPoint)
                    'If xFloorPoint(0) = yFloorPoint(0) And xFloorPoint(1) = yFloorPoint(1) Then
                    'Debug.Print vbTab & yRiser.Name & ": Relative " & yFloorPoint(0) & ", " & yFloorPoint(1) & " Absolute " & yRiser.InsertionPoint.x & ", " & yRiser.InsertionPoint.y
                    '''If yRiser.Name Like "*.L" Then
                        If PointsMatchAlmost(xFloorPoint, yFloorPoint) Then
                            Debug.Print vbTab & "matches " & yRiser.Name
                        End If
                    '''End If '****** test only
                Next yRiser
            '''End If '************* test only
        Next xRiser
    Else
        MsgBox "There are no risers to align"
    End If
End Sub

Public Sub ChangeHead()
Dim pfSS As AcadSelectionSet  'pickfirst
Dim ssObject As IntelliCAD.entity
Dim msg As String
Dim newLabel As String
Dim Block As IntelliCAD.blockInsert
Dim strHeadNumber As String
    msg = vbCrLf
    
    Set pfSS = ActiveDocument.ActiveSelectionSet 'PickfirstSelectionSet
    For Each ssObject In pfSS
        If UCase(ssObject.layer) = UCase("HeadLabels") Then
            'MsgBox "You picked a head label!"
            If ssObject.EntityName = "BlockInsert" Then
                Set Block = ssObject
                strHeadNumber = GetAttrib(ssObject, "HEADNUMBER")
                newLabel = InputBox("What would you like to label this head?", "ChangeHead", strHeadNumber)
                If newLabel <> "" Then
                    SetAttrib ssObject, "HEADNUMBER", newLabel
                End If
            End If
        End If
    Next ssObject
    'MsgBox "The Pickfirst selection set contains: " & msg
    
    

End Sub

Public Function ChangePrefix(OldString As String, NewPrefix As String)
    Dim newString As String
    If Len(OldString) > 3 Then
        If Mid(OldString, 2, 1) = "." Then
            newString = Mid(OldString, 3)
            newString = NewPrefix & newString
        Else
            Err.Raise 1, "ChangePrefix", "The string is missing a dot at position 3 " & OldString
        End If
    Else
        Err.Raise 1, "ChangePrefix", "The string is not long enough " & OldString
    End If
    
    ChangePrefix = newString
    
End Function

Public Function CollectionContainsName(InputCollection As Collection, Match As Variant) As Boolean
Dim bResult As Boolean
Dim vItem As Variant
    bResult = False
    
    For Each vItem In InputCollection
        If vItem.Name = Match.Name Then
            bResult = True
            Exit For
        End If
    Next vItem
    
    CollectionContainsName = bResult
    
End Function
Public Function CollectionContainsNameString(InputCollection As Collection, MatchName As String) As Boolean
Dim bResult As Boolean
Dim vItem As Variant
    bResult = False
    
    For Each vItem In InputCollection
        If vItem.Name = MatchName Then
            bResult = True
            Exit For
        End If
    Next vItem
    
    CollectionContainsNameString = bResult
    
End Function

Public Function ConvertToHead(InputNode As Node) As Head
    Dim newHead As New Head
    
    If InputNode.NodeTypeID = NODE_TYPE_HEAD Then
        newHead.elevation = InputNode.elevation
        Set newHead.insertionPoint = InputNode.insertionPoint
        newHead.model = InputNode.model
        newHead.Name = InputNode.Name
        newHead.NodeNumber = InputNode.NodeNumber
        Set ConvertToHead = newHead
    Else
        Set ConvertToHead = Null
    End If
    
End Function


Public Function ConvertToTee(InputNode As Node) As Tee
    Dim newTee As New Tee
    
    If InputNode.NodeTypeID = NODE_TYPE_TEE Then
        newTee.elevation = InputNode.elevation
        Set newTee.insertionPoint = InputNode.insertionPoint
        newTee.model = InputNode.model
        newTee.Name = InputNode.Name
        newTee.NodeNumber = InputNode.NodeNumber
        Set ConvertToTee = newTee
    Else
        Set ConvertToTee = Null
    End If
    
End Function
Public Function ConvertToFixture(InputNode As Node) As Fixture
    Dim newFixture As New Fixture
    
    If InputNode.NodeTypeID = NODE_TYPE_FIXTURE Then
        newFixture.elevation = InputNode.elevation
        newFixture.insertionPoint = InputNode.insertionPoint
        newFixture.model = InputNode.model
        newFixture.Name = InputNode.Name
        newFixture.NodeNumber = InputNode.NodeNumber
        Set ConvertToFixture = newFixture
    Else
        Set ConvertToFixture = Null
    End If
    
End Function
Public Function ConvertToRiser(InputNode As Node) As Riser
    Dim newRiser As New Riser
    
    If InputNode.NodeTypeID = NODE_TYPE_RISER Then
        newRiser.elevation = InputNode.elevation
        Set newRiser.insertionPoint = InputNode.insertionPoint
        newRiser.Name = InputNode.Name
        newRiser.NodeNumber = InputNode.NodeNumber
        Set ConvertToRiser = newRiser
    Else
        Set ConvertToRiser = Nothing
    End If
    
End Function
Public Function ConvertToManifold(InputNode As Node) As Manifold
    Dim newManifold As New Manifold
    
    If InputNode.NodeTypeID = NODE_TYPE_MANIFOLD_PORT Then
        newManifold.elevation = InputNode.elevation
        Set newManifold.insertionPoint = InputNode.insertionPoint
        newManifold.Name = InputNode.Name
        newManifold.NodeNumber = InputNode.NodeNumber
        Set ConvertToManifold = newManifold
    Else
        Set ConvertToManifold = Nothing
    End If
    
End Function
Public Sub DeleteHeadLabels()
Dim deleteLabels As New Collection
Dim lngObjectIndex As Long
Dim xObject As AcadObject
Dim xHeadLabel As IntelliCAD.blockInsert
    
    For lngObjectIndex = 0 To ActiveDocument.ModelSpace.Count - 1
        Set xObject = ActiveDocument.ModelSpace.Item(lngObjectIndex)
    
        If TypeOf xObject Is IntelliCAD.blockInsert Then
            Set xHeadLabel = xObject
            If UCase(xHeadLabel.layer) = UCase("HeadLabels") _
            Or UCase(xHeadLabel.Name) = UCase("HeadLabel") _
            Or UCase(xHeadLabel.layer) = UCase("0b__Head_IDs") _
            Or UCase(xHeadLabel.Name) = UCase("HeadID") Then deleteLabels.Add xHeadLabel
        End If
    Next lngObjectIndex
    
    For Each xHeadLabel In deleteLabels
        xHeadLabel.Delete
    Next xHeadLabel
    
End Sub
Public Sub DeleteTeeLabels()
Dim deleteLabels As New Collection
Dim lngObjectIndex As Long
Dim xObject As AcadObject
Dim xTeeLabel As IntelliCAD.blockInsert
    
    For lngObjectIndex = 0 To ActiveDocument.ModelSpace.Count - 1
        Set xObject = ActiveDocument.ModelSpace.Item(lngObjectIndex)
    
        If TypeOf xObject Is IntelliCAD.blockInsert Then
            Set xTeeLabel = xObject
            If UCase(xTeeLabel.layer) = UCase("TeeLabels") Or UCase(xTeeLabel.Name) = UCase("TeeLabel") Then deleteLabels.Add xTeeLabel
        End If
    Next lngObjectIndex
    
    For Each xTeeLabel In deleteLabels
        xTeeLabel.Delete
    Next xTeeLabel
    
End Sub
Public Sub DeleteRiserLabels()
Dim deleteLabels As New Collection
Dim lngObjectIndex As Long
Dim xObject As AcadObject
Dim xRiserLabel As IntelliCAD.blockInsert
    
    For lngObjectIndex = 0 To ActiveDocument.ModelSpace.Count - 1
        Set xObject = ActiveDocument.ModelSpace.Item(lngObjectIndex)
    
        If TypeOf xObject Is IntelliCAD.blockInsert Then
            Set xRiserLabel = xObject
            If UCase(xRiserLabel.layer) = UCase("RiserLabels") Or UCase(xRiserLabel.Name) = UCase("RiserLabel") Then deleteLabels.Add xRiserLabel
        End If
    Next lngObjectIndex
    
    For Each xRiserLabel In deleteLabels
        xRiserLabel.Delete
    Next xRiserLabel
    
End Sub


Public Sub DrawNewPipe(OldPipe As IntelliCAD.lwPolyline, PipeVertices As Variant, CurrentVertex As Variant, PipeWidth As Double)
Dim newPipe As IntelliCAD.lwPolyline
Dim verticeCount As Long
Dim verticeIndex As Long
Dim layer As String
    
    layer = OldPipe.layer
        
    CheckAddLayer layer, Colors.vicRed
    
    Set newPipe = ActiveDocument.ModelSpace.AddLightWeightPolyline(PipeVertices)
    
    newPipe.layer = layer
    newPipe.Color = OldPipe.Color
    newPipe.LineType = OldPipe.LineType
    newPipe.LineWeight = OldPipe.LineWeight
    
    verticeCount = ((newPipe.Coordinates.Count - 1) + 1) / 2
    For verticeIndex = 0 To verticeCount - 1
        newPipe.SetWidth verticeIndex, PipeWidth, PipeWidth
    Next verticeIndex
    
    ReDim Preserve PipeVertices(0 To 1)
    PipeVertices(0) = CurrentVertex(0)
    PipeVertices(1) = CurrentVertex(1)
    
End Sub

Public Sub DrawOneLabel(x As Double, y As Double, labelText As String, Color As Long)
Dim t As IntelliCAD.text
Dim p(0 To 2) As Double
Dim textSize As Double
Dim pipeLabelsLayer As String

    pipeLabelsLayer = "Pipe Labels"
    textSize = 4
    
    p(0) = x
    p(1) = y
    p(2) = 0

    CheckAddLayer pipeLabelsLayer, Color, False, True
    
    Set t = ActiveDocument.ModelSpace.AddText(labelText, p, textSize)
    
    t.layer = pipeLabelsLayer
    
    
End Sub

Public Sub DrawPipeLabel(InputPipe As Pipe)
Dim lngVertexCount As Long
Dim iXa As Double
Dim iYa As Double
Dim iXb As Double
Dim iYb As Double
Dim iXm As Double
Dim iYm As Double
Dim i As Long
Dim midPoint As Variant

    CheckAddLayer "Pipe Labels", Colors.vicCyan
    
    If InputPipe.polyline Is Nothing Then
        lngVertexCount = 0
    Else
        lngVertexCount = (InputPipe.polyline.Coordinates.Count - 1) + 1
        If lngVertexCount >= 4 Then
            iXa = 0
            iXa = 0
            For i = 0 To (InputPipe.polyline.Coordinates.Count - 1) Step 4
                iXa = InputPipe.polyline.Coordinates(i)
                iYa = InputPipe.polyline.Coordinates(i + 1)
                If i + 2 <= (InputPipe.polyline.Coordinates.Count - 1) Then
                    iXb = InputPipe.polyline.Coordinates(i + 2)
                    If i + 3 <= (InputPipe.polyline.Coordinates.Count - 1) Then
                        iYb = InputPipe.polyline.Coordinates(i + 3)
                        midPoint = GetMidPoint(iXa, iYa, iXb, iYb)
                        If (midPoint.Count - 1) = 2 Then
                            iXm = midPoint(0)
                            iYm = midPoint(1)
                            ''2011-06-16 MLD'''Debug.Print vbTab & iXa & ", " & iYa & " > " & iXb & ", " & iYb & " MidPoint = " & iXm & ", " & iYm
                            DrawOneLabel iXm, iYm, InputPipe.Name, InputPipe.polyline.Color
                        End If
                    Else
                        MsgBox "This polyline does not have enough points!"
                    End If
                Else
                    ''''Debug.Print vbTab & "End Point"
                End If

                
            Next i
        Else
            MsgBox "This polyline does not have enough points!"
        End If
    End If
    
    ''2011-06-16 MLD'''Debug.Print "Pipe " & InputPipe.Name & " " & lngVertexCount
    
End Sub

Public Sub LabelPipes()
Dim xPipe As Pipe
    
    Debug.Print "Labeling " & Pipes.Count & " pipes"
    For Each xPipe In Pipes
        DrawPipeLabel xPipe
    Next xPipe
    
End Sub

Public Sub DrawTraceLine(Ax As Double, Ay As Double, Bx As Double, By As Double, XOffset As Double, YOffset As Double, LineColor As Integer)
Dim startPoint(0 To 2) As Double
Dim endPoint(0 To 2) As Double
Dim traceLine As IntelliCAD.line

    startPoint(0) = Ax + XOffset
    startPoint(1) = Ay + YOffset
    startPoint(2) = 0
    endPoint(0) = Bx + XOffset
    endPoint(1) = By + YOffset
    endPoint(2) = 0
    Set traceLine = ActiveDocument.ModelSpace.AddLine(startPoint, endPoint)
    traceLine.Color = LineColor
    traceLine.layer = "TEST"
    
End Sub
'Public Function EntityElevation(entity As AcadObject) As Variant
Public Function EntityElevation(entity As IntelliCAD.entity) As Variant
Dim Entit As AcadObject
Dim Boxes As New Collection
Dim Labels As New Collection
Dim Box As IntelliCAD.polyline
Dim eBox As ElevationBox
Dim Coordinates As Variant
Dim ElevLabel As IntelliCAD.text


    'See which Boxes exist
    For Each Entit In ActiveDocument.ModelSpace
        If UCase(Entit.layer) = UCase("ElevationBox") Then
            Debug.Print "Elevation Box Layer!"
            If Entit.EntityName = "Polyline" Then
                Set Box = Entit
                Coordinates = Box.Coordinates
                If (Coordinates.Count - 1) = 7 Then ' an upper bound of 7 means it's a rectangle
                    Debug.Print "Box!"
                    Set eBox = New ElevationBox
                    eBox.Box = Box
                    Boxes.Add eBox
                End If
            ElseIf Entit.EntityName = "Text" Then
                Set ElevLabel = Entit
                Labels.Add ElevLabel
            End If
        End If
    Next Entit
    
    'get elevation text
    For Each eBox In Boxes
        For Each ElevLabel In Labels
            If InBox(eBox.Box, ElevLabel.insertionPoint) Then
                Debug.Print "Elevation label is: " & ElevLabel.TextString
                
                eBox.elevation = 1
            End If
        Next ElevLabel
    Next Entit
    
    'See which Box Entity is in
    For Each eBox In Boxes
        If InBox(eBox.Box, entity.insertionPoint) Then
            EntityElevation = eBox.elevation
            Exit For
        End If
    Next Entit

End Function

Public Sub GetLabels()
    Dim selset As AcadSelectionSet
    Dim entity As IntelliCAD.entity
    
    Set selset = ActiveDocument.SelectionSets.Add("GET_LABELS")
    
    selset.SelectOnScreen
    
    For Each entity In selset
        Debug.Print entity.EntityName
    Next entity
End Sub


Public Sub MatchEntities()
    Dim xPipe As Pipe
    Dim xNodeID As NodeID
    Dim xNode As Node
    Dim dblErrorElevation As Double
    Dim labelHasNode As Boolean
    Dim xTee As Tee
    Dim teeIndex As Long
    
    'Begin Entity Matching *********************************************************************************************
    'Debug.Print "Nodes.Count = " & Nodes.Count
    
    Set Risers = New Collection
    
    For Each xNode In nodes
        'Figure out the elevation...
        xNode.elevation = PointElevation(xNode.insertionPoint, ElevationBoxes)
        
        'NodeID Label to Node Matcher If Block
        'yes they do '''''If xNode.NodeTypeID = NODE_TYPE_RISER Then
            'Risers do not have NodeID labels
            'yes they do '''''Risers.Add ConvertToRiser(xNode)
        'Non Risers: Heads, Tees
        'yes they do '''''Else
            'Match each Node to a NodeID label
            For Each xNodeID In NodeIDs
                If xNodeID.insertionPoint.y = xNode.insertionPoint.y _
                    And xNodeID.insertionPoint.Z = xNode.insertionPoint.Z _
                    And xNodeID.insertionPoint.x = xNode.insertionPoint.x Then
                    xNode.Name = xNodeID.Name
                    If xNode.NodeTypeID = NODE_TYPE_HEAD Then
                        'TODO: newHead.Model = Entit.Name
                        If Not CollectionContainsName(Heads, xNode) Then
                            Heads.Add ConvertToHead(xNode), xNodeID.Name
                        End If
                        
                    ElseIf xNode.NodeTypeID = NODE_TYPE_TEE Then
                        If Not CollectionContainsName(Tees, xNode) Then
                            Tees.Add ConvertToTee(xNode)
                        End If
                    
                    ElseIf xNode.NodeTypeID = NODE_TYPE_FIXTURE Then
                        If Not CollectionContainsName(Fixtures, xNode) Then
                            'Tees.Add ConvertToFixture(xNode)
                            Fixtures.Add ConvertToFixture(xNode)
                        End If
                    
                    ElseIf xNode.NodeTypeID = NODE_TYPE_RISER Then
                        If Not CollectionContainsName(Risers, xNode) Then
                            Risers.Add ConvertToRiser(xNode)
                        End If
                    
                    ElseIf xNode.NodeTypeID = NODE_TYPE_MANIFOLD_PORT Then
                        If Not CollectionContainsName(ManifoldNodes, xNode) Then
                            ManifoldNodes.Add ConvertToManifold(xNode)
                        End If
                    
                    End If
                End If
            Next xNodeID
        'yes they do '''''End If
            
        
        'Match Nodes with pipe ends
        For Each xPipe In Pipes
            If xNode.insertionPoint.x = xPipe.startPoint(0) _
                And xNode.insertionPoint.y = xPipe.startPoint(1) Then
                    'Debug.Print "Pipe Match " & xPipe.PipeNumber & " " & xPipe.ActualLength & " Node: " & xNode.Name
                    Set xPipe.StartNode = xNode
            ElseIf PointsMatchAlmost(xNode.insertionPoint, xPipe.startPoint, 0.01) Then
                Debug.Print "Pipe Point Match Almost " & xPipe.PipeNumber & " " & xPipe.ActualLength & " Node: " & xNode.Name
                '2014'ErrorCircle xNode.InsertionPoint, 130, "Almost " & xPipe.Name, 0.111, 0.211, 0.311
            End If
            If xNode.insertionPoint.x = xPipe.endPoint(0) _
                And xNode.insertionPoint.y = xPipe.endPoint(1) Then
                    'Debug.Print "Pipe Match " & xPipe.PipeNumber & " " & xPipe.ActualLength & " Node: " & xNode.Name
                    Set xPipe.EndNode = xNode
            ElseIf PointsMatchAlmost(xNode.insertionPoint, xPipe.endPoint, 0.01) Then
                Debug.Print "Pipe Point Match Almost " & xPipe.PipeNumber & " " & xPipe.ActualLength & " Node: " & xNode.Name
                '2014'ErrorCircle xNode.InsertionPoint, 130, "Almost " & xPipe.Name, 0.111, 0.211, 0.311
            End If
            
        Next xPipe
        
        'Error in node name
        If Trim(xNode.Name) = "" Then
            dblErrorElevation = PointElevation(xNode.insertionPoint)
            Debug.Print "Missing Name at " & xNode.insertionPoint.x & ", " & xNode.insertionPoint.y & ", " & xNode.insertionPoint.Z & ", Elevation = " & dblErrorElevation & " make sure the insertion point of the Head or Tee label lines up with the head or tee itself."
            '2014'ErrorCircle xNode.InsertionPoint, Colors.vicCyan, "Missing Name"
        End If
    
    Next xNode
    
    Debug.Print "***Checking Node Labels"
    For Each xNodeID In NodeIDs
        labelHasNode = False
        For Each xNode In nodes
            If xNodeID.insertionPoint.x = xNode.insertionPoint.x And xNodeID.insertionPoint.y = xNode.insertionPoint.y Then
                labelHasNode = True
                Exit For
            End If
        Next xNode
        If Not labelHasNode Then
            dblErrorElevation = PointElevation(xNodeID.insertionPoint)
            Debug.Print "Label missing a node at " & xNodeID.insertionPoint.x & ", " & xNodeID.insertionPoint.y & ", " & xNodeID.insertionPoint.Z & ", Elevation = " & dblErrorElevation & " make sure the insertion point of the Head or Tee label lines up with the head or tee itself."
            '2014'ErrorCircle xNodeID.InsertionPoint, Colors.vicCyan, "Missing Node @ " & xNodeID.Name
        End If
    Next xNodeID
    
    Dim bNodeHasAPipe As Boolean
    
    For Each xNode In nodes
        bNodeHasAPipe = False
        For Each xPipe In Pipes
            
            If PointsMatch(xPipe.startPoint, xNode.insertionPoint) _
                Or PointsMatch(xPipe.endPoint, xNode.insertionPoint) Then
                bNodeHasAPipe = True
                Exit For
            Else
            End If
        Next xPipe
        If Not bNodeHasAPipe Then
            '2014'ErrorCircle xNode.InsertionPoint, vicYellow, "No Pipes!", 10.8, 20.8, 60.8
        End If
    Next xNode
    
    
    teeIndex = 1
    For Each xTee In Tees
        If xTee.Name = "S.0" Then
            Tees.Remove teeIndex
        End If
        teeIndex = teeIndex + 1
    Next xTee
    
    'Change any tee named "S.0" to a manifold port
    For Each xNode In nodes
        If xNode.Name = "S.0" And xNode.NodeTypeID <> NODE_TYPE_MANIFOLD_PORT Then
            xNode.NodeTypeID = NODE_TYPE_MANIFOLD_PORT
            If Not CollectionContainsName(ManifoldNodes, xNode) Then
                ManifoldNodes.Add ConvertToManifold(xNode)
            End If
        End If
    Next xNode
    
End Sub
Public Function FloorCoordinates(InputPoint As Variant) As Variant
Dim closestFloorTag As FloorTag
Dim relative As Variant

    Set closestFloorTag = GetElevationBoxFloorTag(InputPoint)
    
    If closestFloorTag Is Nothing Then Exit Function
    
    relative = RelativeCoordinates(closestFloorTag.insertionPoint, InputPoint)
    
    FloorCoordinates = relative

End Function

Public Function GetMidPoint(startX As Double, startY As Double, endX As Double, endY As Double) As Variant
    Dim vReturn As Variant
    Dim dblResult(0 To 2) As Double
    
    dblResult(0) = (startX + endX) / 2
    dblResult(1) = (startY + endY) / 2
    dblResult(2) = 0
    
    GetMidPoint = dblResult
        
End Function

Public Sub CircleLabelBlock(InputPoint As Variant, Label As String, Optional InputColor As Variant, Optional CircleSize As Variant, Optional TextRotation As Variant)
Dim newCircle As IntelliCAD.Circle
Dim hlabel As IntelliCAD.text

    If IsMissing(InputColor) Then InputColor = Colors.vicGreen
    If IsMissing(CircleSize) Then CircleSize = 0.2
    If IsMissing(TextRotation) Then TextRotation = 0
    
    Set newCircle = ActiveDocument.ModelSpace.AddCircle(InputPoint, CircleSize)
    newCircle.layer = "TEST"
    newCircle.Color = InputColor
    
    If Label <> "" Then
        Set hlabel = ActiveDocument.ModelSpace.AddText(Label, InputPoint, 0.18)
        hlabel.Rotation = TextRotation
        hlabel.layer = "TEST"
        hlabel.Color = InputColor
    End If
    
End Sub

Public Sub LabelNodes()
Dim lastNodeNumber As Long

    ScanEntities
    DeleteHeadLabels
    DeleteTeeLabels
    DeleteRiserLabels
    'LabelHeadsTees
    'lastNodeNumber = LabelHeads()
    'lastNodeNumber = LabelTees(lastNodeNumber + 1)
    lastNodeNumber = 1
    lastNodeNumber = LabelNodesByType(NODE_TYPE_HEAD, "HeadLabels", "HeadLabel.dwg", "HEADNUMBER", "H.", lastNodeNumber)
    lastNodeNumber = lastNodeNumber + 1
    lastNodeNumber = LabelNodesByType(NODE_TYPE_TEE, "TeeLabels", "TeeLabel.dwg", "TEENUMBER", "T.", lastNodeNumber)
    lastNodeNumber = lastNodeNumber + 1
    lastNodeNumber = LabelNodesByType(NODE_TYPE_RISER, "RiserLabels", "RiserLabel.dwg", "RISERNUMBER", "R.", lastNodeNumber)
    
    Debug.Print "Last Node Number was " & lastNodeNumber
    
End Sub
Public Sub LabelHeadsTees()
    Dim xPipe As Pipe
    Dim xNodeID As NodeID
    Dim xNode As Node
    Dim dblErrorElevation As Double
    Dim newNodeID As NodeID
    Dim Block As IntelliCAD.blockInsert
    Dim indexHead As Long
    Dim indexTee As Long
    Dim bHasTeeLabelLayer As Boolean
    Dim bHasHeadLabelLayer As Boolean
    Dim lay As IntelliCAD.layer
    Dim layerheadlabels As IntelliCAD.layer
    Dim layerTeeLabels As IntelliCAD.layer
    
  
    Set Risers = New Collection
    
    
    CheckAddLayer "HeadLabels", Colors.vicGreen
    CheckAddLayer "TeeLabels", Colors.vicWhite

    
    
    For Each xNode In nodes
        Debug.Print "Labeling node " & xNode.NodeNumber & "..."
        If xNode.NodeTypeID = NODE_TYPE_HEAD Then
            Set Block = ActiveDocument.ModelSpace.InsertBlock(xNode.insertionPoint, "HeadLabel.dwg", 1, 1, 1, 0)
            Block.layer = "HeadLabels"
            
            indexHead = indexHead + 1
            SetAttrib Block, "HEADNUMBER", "H." & indexHead
            
            
'        ElseIf xNode.NodeTypeID = NODE_TYPE_TEE Then
'            Set block = ActiveDocument.ModelSpace.InsertBlock(xNode.InsertionPoint, "TeeLabel.dwg", 1, 1, 1, 0)
'            block.Layer = "TeeLabels"
'
'            indexTee = indexTee + 1
'            SetAttrib block, "TEENUMBER", "T." & indexTee
                
            
        End If
        
            
        
        'Error in node name
        If Trim(xNode.Name) = "" Then
            '''dblErrorElevation = PointElevation(xNode.InsertionPoint)
            ''''Debug.Print "Missing Name at " & xNode.InsertionPoint.x & ", " & xNode.InsertionPoint.y & ", " & xNode.InsertionPoint.z & ", Elevation = " & dblErrorElevation & " make sure the insertion point of the Head or Tee label lines up with the head or tee itself."
            '''ErrorCircle xNode.InsertionPoint, Colors.vicCyan, "Missing Name"
        End If
    
    Next xNode
    
    indexTee = indexHead
    For Each xNode In nodes
        Debug.Print "Labeling node " & xNode.NodeNumber & "..."
        If xNode.NodeTypeID = NODE_TYPE_TEE Then
            Set Block = ActiveDocument.ModelSpace.InsertBlock(xNode.insertionPoint, "TeeLabel.dwg", 1, 1, 1, 0)
            Block.layer = "TeeLabels"
            
            indexTee = indexTee + 1
            SetAttrib Block, "TEENUMBER", "T." & indexTee
                
            
        End If
        
            
        
        'Error in node name
        If Trim(xNode.Name) = "" Then
            '''dblErrorElevation = PointElevation(xNode.InsertionPoint)
            ''''Debug.Print "Missing Name at " & xNode.InsertionPoint.x & ", " & xNode.InsertionPoint.y & ", " & xNode.InsertionPoint.z & ", Elevation = " & dblErrorElevation & " make sure the insertion point of the Head or Tee label lines up with the head or tee itself."
            '''ErrorCircle xNode.InsertionPoint, Colors.vicCyan, "Missing Name"
        End If
    
    Next xNode
    
    
End Sub


Public Function LabelHeads(Optional StartNodeNumber As Variant) As Long
'Returns last or highest head number
    Dim xNode As Node
    Dim Block As IntelliCAD.blockInsert
    Dim indexHead As Long
    Dim layerheadlabels As IntelliCAD.layer
  
    Set Risers = New Collection
    If IsMissing(StartNodeNumber) Then StartNodeNumber = 1
    indexHead = StartNodeNumber
    CheckAddLayer "HeadLabels", Colors.vicGreen
    
    For Each xNode In nodes
        Debug.Print "Labeling Head " & xNode.NodeNumber & "..."
        If xNode.NodeTypeID = NODE_TYPE_HEAD Then
            Set Block = ActiveDocument.ModelSpace.InsertBlock(xNode.insertionPoint, "HeadLabel.dwg", 1, 1, 1, 0)
            Block.layer = "HeadLabels"
            SetAttrib Block, "HEADNUMBER", "H." & indexHead
            indexHead = indexHead + 1
        End If
    Next xNode
    
    LabelHeads = indexHead
    
End Function

Public Function LabelNodesByType(NodeTypeID As Long, layerName As String, BlockDwgName As String, NumberAttributeName As String, NodeNumberPrefix As String, Optional StartNodeNumber As Variant) As Long
'Returns last or highest head number
    Dim xNode As Node
    Dim Block As IntelliCAD.blockInsert
    Dim indexNode As Long
    Dim layerheadlabels As IntelliCAD.layer
  
    Set Risers = New Collection
    If IsMissing(StartNodeNumber) Then StartNodeNumber = 1
    indexNode = StartNodeNumber
    CheckAddLayer layerName, Colors.vicGreen
    
    
    
    For Each xNode In nodes
        Debug.Print "Labeling Node " & xNode.NodeNumber & "..."
        If xNode.NodeTypeID = NodeTypeID Then
            Set Block = ActiveDocument.ModelSpace.InsertBlock(xNode.insertionPoint, BlockDwgName, 1, 1, 1, 0)
            Block.layer = layerName
            SetAttrib Block, NumberAttributeName, NodeNumberPrefix & indexNode
            xNode.NodeNumber = indexNode
            indexNode = indexNode + 1
        End If
    Next xNode
    
    LabelNodesByType = indexNode
    
End Function
Public Function LabelTees(Optional StartNodeNumber As Variant) As Long
'Returns highest tee node number
    Dim xNode As Node
    Dim Block As IntelliCAD.blockInsert
    Dim indexTee As Long
  
    Set Risers = New Collection
    If IsMissing(StartNodeNumber) Then StartNodeNumber = 1
    indexTee = StartNodeNumber
    CheckAddLayer "TeeLabels", Colors.vicGreen
    
    For Each xNode In nodes
        Debug.Print "Labeling Tee " & xNode.NodeNumber & "..."
        If xNode.NodeTypeID = NODE_TYPE_TEE Then
            Set Block = ActiveDocument.ModelSpace.InsertBlock(xNode.insertionPoint, "TeeLabel.dwg", 1, 1, 1, 0)
            Block.layer = "TeeLabels"
            SetAttrib Block, "TEENUMBER", "T." & indexTee
            indexTee = indexTee + 1
        End If
    Next xNode
    
    LabelTees = indexTee
    
End Function

Private Sub AcadActiveDocument_SelectionChanged()  'this only works in "ThisActiveDocument" I think... not checked.
  Dim pfSS As AcadSelectionSet  'pickfirst
    Dim ssObject As IntelliCAD.entity

    Dim msg As String
    msg = vbCrLf
    
    Set pfSS = Document.PickfirstSelectionSet
    For Each ssObject In pfSS
        'msg = msg & vbCrLf & ssobject.EntityName
        If UCase(ssObject.layer) = UCase("HeadLabels") Then
            'MsgBox "You picked a head label!"
        End If
    Next ssObject
    'MsgBox "The Pickfirst selection set contains: " & msg
    
    

End Sub
Public Function ExtractLong(InputString As String) As Long
    Dim ReturnLong As Long
    Dim FirstDigit As Long
    Dim PreviousFoundDigit As Long
    Dim ThisFoundDigit As Long
    
    
    
    ThisFoundDigit = InStr(1, InputString, "0")
    If ThisFoundDigit > 0 And (FirstDigit = 0 Or ThisFoundDigit < FirstDigit) Then FirstDigit = ThisFoundDigit
    ThisFoundDigit = InStr(1, InputString, "1")
    If ThisFoundDigit > 0 And (FirstDigit = 0 Or ThisFoundDigit < FirstDigit) Then FirstDigit = ThisFoundDigit
    ThisFoundDigit = InStr(1, InputString, "2")
    If ThisFoundDigit > 0 And (FirstDigit = 0 Or ThisFoundDigit < FirstDigit) Then FirstDigit = ThisFoundDigit
    ThisFoundDigit = InStr(1, InputString, "3")
    If ThisFoundDigit > 0 And (FirstDigit = 0 Or ThisFoundDigit < FirstDigit) Then FirstDigit = ThisFoundDigit
    ThisFoundDigit = InStr(1, InputString, "4")
    If ThisFoundDigit > 0 And (FirstDigit = 0 Or ThisFoundDigit < FirstDigit) Then FirstDigit = ThisFoundDigit
    ThisFoundDigit = InStr(1, InputString, "5")
    If ThisFoundDigit > 0 And (FirstDigit = 0 Or ThisFoundDigit < FirstDigit) Then FirstDigit = ThisFoundDigit
    ThisFoundDigit = InStr(1, InputString, "6")
    If ThisFoundDigit > 0 And (FirstDigit = 0 Or ThisFoundDigit < FirstDigit) Then FirstDigit = ThisFoundDigit
    ThisFoundDigit = InStr(1, InputString, "7")
    If ThisFoundDigit > 0 And (FirstDigit = 0 Or ThisFoundDigit < FirstDigit) Then FirstDigit = ThisFoundDigit
    ThisFoundDigit = InStr(1, InputString, "8")
    If ThisFoundDigit > 0 And (FirstDigit = 0 Or ThisFoundDigit < FirstDigit) Then FirstDigit = ThisFoundDigit
    ThisFoundDigit = InStr(1, InputString, "9")
    If ThisFoundDigit > 0 And (FirstDigit = 0 Or ThisFoundDigit < FirstDigit) Then FirstDigit = ThisFoundDigit
    
    If FirstDigit > 0 Then
        ReturnLong = CLng(val(Mid(InputString, FirstDigit)))
    Else
        ReturnLong = 0
    End If
    
    ExtractLong = ReturnLong
    
End Function

Public Function GetEndPoint(InputLine As IntelliCAD.lwPolyline) As Variant

    GetEndPoint = InputLine.Coordinate(((InputLine.Coordinates.Count - 1) - 1) / 2)
    
End Function
Public Sub HeadPairMatch()
Dim xHeadPair As HeadPair
Dim xHead As Head
Dim BadHeadPairs As New Collection

    For Each xHeadPair In HeadPairs
        For Each xHead In Heads
            If PointsMatch(xHeadPair.HeadPairLine.startPoint, xHead.insertionPoint) Then
                Set xHeadPair.HeadA = xHead
            End If
            If PointsMatch(xHeadPair.HeadPairLine.endPoint, xHead.insertionPoint) Then
                Set xHeadPair.HeadB = xHead
            End If
        Next xHead
    Next xHeadPair
    
    For Each xHeadPair In HeadPairs
        If xHeadPair.HeadA Is Nothing Then
            BadHeadPairs.Add xHeadPair
            '2014'ErrorCircle xHeadPair.HeadPairLine.startPoint, Colors.vicCyan, "HP", 15, 17, 100
            '2014'ErrorCircle xHeadPair.HeadPairLine.EndPoint, Colors.vicWhite, "HP", 10, 20, 120
        ElseIf xHeadPair.HeadB Is Nothing Then
            BadHeadPairs.Add xHeadPair
            '2014'ErrorCircle xHeadPair.HeadPairLine.EndPoint, Colors.vicCyan, "HP", 15, 17, 100
            '2014'ErrorCircle xHeadPair.HeadPairLine.startPoint, Colors.vicWhite, "HP", 10, 20, 120
        End If
    Next xHeadPair
    
    If BadHeadPairs.Count > 0 Then
        MsgBox "There are " & BadHeadPairs.Count & " head pairs not linked to heads correctly.  " _
            & "A head pair is the same as a '2Hd Calc Connection'", vbCritical, "BadHeadPairs"
    End If
    

End Sub

Public Function GetStartPoint(InputLine As IntelliCAD.lwPolyline) As Variant

    GetStartPoint = InputLine.Coordinate(0)
    
End Function
Public Sub ListNodes(InputCollection As Collection)
    Dim xNode As Node
    For Each xNode In InputCollection
        Debug.Print xNode.NodeNumber & ": " & xNode.Name & ": " & xNode.NodeTypeID
    Next xNode
End Sub

Public Sub ListPipes(InputCollection As Collection)
    Dim xPipe As Pipe
    Dim strStartName As String
    Dim strEndName As String
    
    Debug.Print "Pipe Count = " & InputCollection.Count
    For Each xPipe In InputCollection
        If xPipe.StartNode Is Nothing Then
            strStartName = "Error"
        Else
            strStartName = xPipe.StartNode.Name
        End If
        If xPipe.EndNode Is Nothing Then
            strEndName = "Error"
        Else
            strEndName = xPipe.EndNode.Name
        End If
        Debug.Print xPipe.PipeNumber & ": " & xPipe.Name & " Start: " & strStartName & " End: " & strEndName
    Next xPipe
    
End Sub

Public Sub OpenHeadModelBarLoader()
    frmHeadModelBars.Show
    
End Sub

Public Sub PrintNamesPoints(Col As Collection)
Dim v As Variant

    If Not Col Is Nothing Then
        If Col.Count > 0 Then
            For Each v In Col
                Debug.Print "Name = " & v.Name '& " " & v.InsertionPoint.x & ", " & v.InsertionPoint.y
            Next v
        Else
            Debug.Print "Collection.Count = " & Col.Count
        End If
    Else
        Debug.Print "Collection is nothing"
    End If
End Sub

Public Sub RemoveDuplicatePipes()
    Dim ent As IntelliCAD.entity
    'Dim dupPipes As New Collection
    Dim allPipes As New Collection
    Dim Pipe As IntelliCAD.lwPolyline
    Dim otherPipe As IntelliCAD.lwPolyline
    Dim otherPipes As New Collection
    Dim dupPipes As New Collection
    Dim allDupPipes As New Collection
    Dim dupPipe As IntelliCAD.lwPolyline
    Dim dupPipeIndex As Long
    
    For Each ent In Document.ModelSpace
        If UCase(Mid(ent.layer, 1, 6)) = UCase("Pipes.") _
            Or UCase(ent.layer) = UCase("Pipes") Then
            If ent.EntityName = "Polyline" Then
                Set Pipe = ent
                allPipes.Add Pipe
            End If
        End If
    Next ent
    
    For Each Pipe In allPipes
        otherPipes.Add Pipe
    Next Pipe
    
    For Each Pipe In allPipes
        Set dupPipes = New Collection
        For Each otherPipe In otherPipes
            If CheckPipeDuplicate(Pipe, otherPipe) Then
                dupPipes.Add otherPipe
                
            End If
        Next otherPipe
        If dupPipes.Count > 1 Then
            Debug.Print "Pipe.Count = " & dupPipes.Count
            dupPipeIndex = 0
            For Each dupPipe In dupPipes
                If dupPipeIndex > 0 Then
                    Debug.Print "Duplicate @ " & Pipe.Coordinates(0) & ", " & Pipe.Coordinates(1)
                    allDupPipes.Add dupPipe
                End If
                dupPipeIndex = dupPipeIndex + 1
            Next dupPipe
        End If
    Next Pipe
    
    dupPipeIndex = 0
    Debug.Print "all duplicate pipes. count = " & allDupPipes.Count
    For Each dupPipe In allDupPipes
        Debug.Print "deleting " & dupPipeIndex & " ...."
        On Error Resume Next
        dupPipe.Delete
        dupPipeIndex = dupPipeIndex + 1
    Next dupPipe
    
End Sub

Public Function CheckPipeDuplicate(a As IntelliCAD.lwPolyline, b As IntelliCAD.lwPolyline) As Boolean

Dim v As Long
    If (a.Coordinates.Count - 1) = (b.Coordinates.Count - 1) Then
        For v = 0 To (a.Coordinates.Count - 1) - 1 Step 2
            If a.Coordinates(v) = b.Coordinates(v) _
                And a.Coordinates(v + 1) = b.Coordinates(v + 1) Then
                CheckPipeDuplicate = True
            Else
                CheckPipeDuplicate = False
                Exit Function
            End If
        Next v
    Else
        CheckPipeDuplicate = False
        Exit Function
    End If
End Function

Public Sub ScanAndMatchEntities()
    
    If ActiveDocument.ModelSpace.Count = 0 Then
        Err.Raise 1000, "Scanner.RunScanAndMatchEntities", "There are no entities in Model Space. " _
         & "Probably no drawing file is opened."
    End If
    
    DeleteErrorCircles
    ScanEntities
    ScanElevationBoxes
    MatchEntities
    HeadPairMatch
    FindOrphans
    FindZeroLengthPipe
    
    
End Sub
Public Sub RunScanAndMatchEntities()
    ScanAndMatchEntities
    MsgBox "All entities have been scanned.  There are " & ErrorCircleCount & " errors circled."
End Sub
Public Sub ScanNodeLabels(NodePrefix As String, NewPrefix As String)
    
    Dim nodeLabel As NodeID
    
    If NodeIDs Is Nothing Or NodeIDs.Count < 1 Then
        ScanEntities
        MatchEntities
    End If
    
    For Each nodeLabel In NodeIDs
        If Mid(nodeLabel.Name, 1, 2) = NodePrefix Then
            Debug.Print "nodeLabel " & nodeLabel.Name
        End If
        
    Next nodeLabel
    
End Sub
Public Sub ScanToDB()
Dim NodeIDs As Collection
Dim lngCalculationID As Long
Dim strCalculationName As String
    
    DeleteErrorCircles
    ScanEntities
    ScanElevationBoxes
    MatchEntities
    HeadPairMatch
    FindOrphans
    
    If Not DeleteNodesPipes() Then
        Exit Sub
    Else
       
        strCalculationName = InputBox("What would you like to name this calculation?", "Export Sprinkler to Datbase File", Job.JobName & " " & Now())
        If strCalculationName = "" Then
            MsgBox "You didn't select a name, the exporting has been cancelled."
            Exit Sub
        End If
        lngCalculationID = AddCalculation(strCalculationName)
    End If
    
    basData.AddNodes nodes, "imported", lngCalculationID
        
    basData.AddHeadPairs HeadPairs, lngCalculationID
        
    basData.AddPipes Pipes, lngCalculationID
    
    MsgBox "Sprinker System Scan Complete!  Open the """ & MDB_FILE_NAME & """ from LoopCalc.", vbInformation, "Sprinkler System Scan"
    
End Sub

Public Function PointElevation(point As Variant, Optional InputElevationBoxes As Variant) As Variant
Dim Entit As IntelliCAD.entity
Dim Boxes As New Collection
Dim Labels As New Collection
Dim Box As IntelliCAD.lwPolyline
Dim Box2d As IntelliCAD.polyline
Dim eBox As ElevationBox
Dim innerBox As ElevationBox
Dim Coordinates As Variant
Dim ElevLabel As IntelliCAD.text
Dim ElevMLable As IntelliCAD.mText
Dim boxCorner(0 To 1) As Double
   
    If IsMissing(InputElevationBoxes) Then
        ScanElevationBoxes
    End If
    Set Boxes = ElevationBoxes

    'See which Box Entity is in
    For Each eBox In Boxes
        If InBox(eBox.Box, point) Then
            PointElevation = eBox.elevation
            If Not eBox.InnerBoxes Is Nothing Then
                For Each innerBox In eBox.InnerBoxes
                    If InBox(innerBox.Box, point) Then
                        'This point is in an inner box, which overrides the outer box (EBox)
                        PointElevation = innerBox.elevation
                        Exit For
                    End If
                Next innerBox
            End If
            Exit For
        End If
    Next eBox

End Function





Public Function PointsMatch(startPoint As Variant, endPoint As Variant) As Boolean
Dim boolResult As Boolean
boolResult = False
    If IsArray(endPoint) And IsArray(startPoint) Then
        If (endPoint.Count - 1) = 2 And (startPoint.Count - 1) = 2 Then
            If startPoint(0) = endPoint(0) _
                And startPoint(1) = endPoint(1) _
                And startPoint(2) = endPoint(2) Then
                PointsMatch = True
            End If
        ElseIf ((endPoint.Count - 1) = 1 Or (endPoint.Count - 1) = 2) _
            And ((startPoint.Count - 1) = 1 Or (startPoint.Count - 1) = 2) Then
            If startPoint(0) = endPoint(0) _
                And startPoint(1) = endPoint(1) Then
                PointsMatch = True
            End If
        End If
    End If
End Function
Public Function PointsMatchAlmost(startPoint As Variant, endPoint As Variant, Optional closeness As Double) As Boolean
Dim boolResult As Boolean
If IsMissing(closeness) Then closeness = POINTS_MATCH_ALMOST
boolResult = False
    If IsArray(endPoint) And IsArray(startPoint) Then
        If (endPoint.Count - 1) = 2 And (startPoint.Count - 1) = 2 Then
            If Abs(startPoint(0) - endPoint(0)) < closeness _
                And Abs(startPoint(1) - endPoint(1)) < closeness _
                And Abs(startPoint(2) - endPoint(2)) < closeness Then
                PointsMatchAlmost = True
            End If
        ElseIf ((endPoint.Count - 1) = 1 Or (endPoint.Count - 1) = 2) _
            And ((startPoint.Count - 1) = 1 Or (startPoint.Count - 1) = 2) Then
            If startPoint(0) = endPoint(0) _
                And Abs(startPoint(1) - endPoint(1)) < closeness Then
                PointsMatchAlmost = True
            End If
        End If
    End If
End Function
Public Function ReSortByName(InputCollection As Collection) As Collection
Dim ReturnCol As New Collection
Dim lngIndex As Long
Dim lngLowestIndex As Long
Dim lngHighestIndex As Long
Dim nextIndex As Long
Dim x As Long
Dim v As Variant
lngIndex = 0
lngLowestIndex = 0
    For Each v In InputCollection
        lngIndex = ExtractLong(v.Name)
        If lngIndex < lngLowestIndex Then lngLowestIndex = lngIndex
        If lngIndex > lngHighestIndex Then lngHighestIndex = lngIndex
    Next v
    
    nextIndex = lngLowestIndex
    While nextIndex <= lngHighestIndex
        For Each v In InputCollection
            lngIndex = ExtractLong(v.Name)
            If lngIndex = nextIndex Then
                ReturnCol.Add v
                Exit For
            End If
        Next v
        nextIndex = nextIndex + 1
    Wend
    
    If InputCollection.Count <> ReturnCol.Count Then
        MsgBox "Error! in ReSortByName", vbCritical
        Exit Function
        Resume
    End If
    Set ReSortByName = ReturnCol
End Function

Public Function PipeDiameterOfLineType(LineType As String) As Double
    Select Case LineType
        Case "Pex-1"
            PipeDiameterOfLineType = 0.86
        Case "Pex- 1h"
            PipeDiameterOfLineType = 0.485
        Case "Pex- 3q"
            PipeDiameterOfLineType = 0.675
        Case Else
            PipeDiameterOfLineType = 0
    End Select
End Function
Public Function PipeDiameterOfColor(LineColor As Long, Optional layer As Variant) As Double
Dim LayerColor As Long
    If IsMissing(layer) Then
        LayerColor = 0
    Else
        LayerColor = Document.Layers.Item(layer).Color
    End If
    
    
    Select Case LineColor
        Case Colors.vicCyan '1 inch
            PipeDiameterOfColor = 0.86
        Case Colors.vicBlue '1 inch
            PipeDiameterOfColor = 0.86
        Case 150    '1 inch (blue)
            PipeDiameterOfColor = 0.86
        Case Colors.vicRed '3/4 inch
            PipeDiameterOfColor = 0.675
        Case Colors.vicGreen '1/2 inch
            PipeDiameterOfColor = 0.485
        Case Colors.vicMagenta '1-1/4 inch
            PipeDiameterOfColor = 1.07
        Case vicYellow '1-1/2 inch
            PipeDiameterOfColor = 1.263
        Case Colors.vicByLayer
            If LayerColor <> 0 Then
                PipeDiameterOfColor = PipeDiameterOfColor(LayerColor)
            Else
                PipeDiameterOfColor = 0
            End If
        Case Else
            PipeDiameterOfColor = 0
    End Select
End Function
Public Function Measure(polyline As IntelliCAD.entity) As Double 'Coordinates As Variant) As Double
'Public Function Measure(polyline As AcadObject) As Double 'Coordinates As Variant) As Double
Dim acc As Double
Dim i As Long
Dim x As Long
Dim y As Long
Dim Coordinates As Variant
    
    If polyline.EntityName = "Polyline" Then
        Coordinates = polyline.Coordinates
    
        acc = 0
        For i = 0 To (Coordinates.Count - 1) Step 2
            
            acc = acc + Math.Sqr((Coordinates(i) - Coordinates(i + 2)) ^ 2 + (Coordinates(i + 1) - Coordinates(i + 3)) ^ 2)
            If (Coordinates.Count - 1) = i + 3 Then Exit For
        
        Next i
    
        Measure = acc
    Else
        Measure = 0
    End If
    

End Function
Public Function InBox(polyline As IntelliCAD.entity, XY As Variant) As Boolean
'Public Function InBox(polyline As AcadObject, XY As Variant) As Boolean
Dim acc As Double
Dim i As Long
Dim XMax As Double
Dim XMin As Double
Dim YMax As Double
Dim YMin As Double
Dim Coordinates As Variant
    
    Dim x As Double
    Dim y As Double
    
    x = XY(0)
    y = XY(1)
    
    If polyline.EntityName = "Polyline" Then
        Coordinates = polyline.Coordinates
        'Debug.Print ".Count - 1 Coor = " & (Coordinates.Count - 1)
        acc = 0
        ''For i = 0 To (Coordinates.Count - 1) Step 2
            
            'Debug.Print "x = " & Coordinates(i) & " Y= " & Coordinates(i + 1)
       
        ''Next i
        
        If (Coordinates.Count - 1) = 7 Then 'rectangle
            XMin = Coordinates(0)
            YMax = Coordinates(1)
            XMax = Coordinates(4)
            YMin = Coordinates(5)
            For i = 0 To (Coordinates.Count - 1) Step 2 'Get X Min
                If XMin > Coordinates(i) Then XMin = Coordinates(i)
                If XMax < Coordinates(i) Then XMax = Coordinates(i)
            Next i
            For i = 1 To (Coordinates.Count - 1) Step 2 'Get Y Min
                If YMin > Coordinates(i) Then YMin = Coordinates(i)
                If YMax < Coordinates(i) Then YMax = Coordinates(i)
            Next i
            If x <= XMax And x >= XMin And y <= YMax And y >= YMin Then
                InBox = True
            Else
                InBox = False
            End If
        End If
    
        'Measure = acc
    Else
        'Measure = 0
    End If
    

End Function

Public Function MeasureAllPipe2() As Long

    Debug.Print "Layers.Count = " & ActiveDocument.Layers.Count
    Dim lay As IntelliCAD.layer 'AutoCAD.AcadLayer
    Dim myLay As IntelliCAD.layer 'AutoCAD.AcadLayer
    Dim newHeadID As HeadID
    Dim newHead As Head
    Dim newPipe As Pipe
    Dim newHeadIDPoint As Variant
    Dim lngLastCoord As Long
    Dim pipes75Length As Double
    Dim pipes1Length As Double
    
    
    Set nodes = New Collection
    Set Heads = New Collection
    Set Pipes = New Collection
    
    
    Dim allPipesLength As Double, pipeLength As Double, headIndex As Long, riserIndex As Long
    allPipesLength = 0: entitIndex = 1: headIndex = 1: riserIndex = 1
    pipes75Length = 0: pipes1Length = 0
    
    Dim Entit As IntelliCAD.entity 'AutoCAD.Entity
    For Each Entit In ActiveDocument.ModelSpace
        'Pipes
        If Entit.EntityName = "Polyline" Then
            If UCase(Entit.layer) = UCase("0e__Pipes-R-FP") Then
                Set newPipe = New Pipe
                
                newPipe.Diameter = PipeDiameterOfColor(Entit.Color, Entit.layer)
                
                pipeLength = Measure(Entit)
                allPipesLength = allPipesLength + pipeLength
                newPipe.ActualLength = pipeLength
                If newPipe.Diameter = 0.675 Then
                    pipes75Length = pipes75Length + pipeLength
                ElseIf newPipe.Diameter = 0.86 Then
                    pipes1Length = pipes1Length + pipeLength
                Else
                    Debug.Print "This pipe.Diameter is not handled " & newPipe.Diameter
                End If
                
                newPipe.PipeNumber = entitIndex
                lngLastCoord = ((Entit.Coordinates.Count - 1) - 1) / 2
                newPipe.endPoint = Entit.Coordinate(lngLastCoord)
                newPipe.startPoint = Entit.Coordinate(0)
                Pipes.Add newPipe
                Debug.Print "Pipe " & entitIndex & vbTab & ".L=" & Format(pipeLength, "0.00 in") & " " & vbTab & newPipe.Diameter
                entitIndex = entitIndex + 1
            End If
        End If
        
        'Heads & HeadIDs
        If Entit.EntityName = "BlockInsert" Then
            'RFC431616-155
            
            'Fittings, Risers
            If Entit.Name = "PipeFitting" _
                Or Entit.Name = "vpilc" Then
                Set newHead = New Head
                newHead.insertionPoint = Entit.insertionPoint
                newHead.NodeNumber = headIndex
                If Entit.Name = "vpilc" Then
                    newHead.Name = "R." & riserIndex
                    riserIndex = riserIndex + 1
                End If
                Heads.Add newHead
            
            headIndex = headIndex + 1
            
            
            'Heads
            ElseIf UCase(Entit.layer) = UCase("0a__Heads") Then
                Debug.Print "Head " & headIndex & " " & Entit.Name
                
                'If we have a head, add it to the Heads Collection
                '...but not if it's the manifold
                If UCase(Entit.Name) <> UCase("Manifold_Info") Then
                    Set newHead = New Head
                    newHead.insertionPoint = Entit.insertionPoint
                    newHead.NodeNumber = headIndex
                    newHead.model = Entit.Name
                    Heads.Add newHead
                End If
                
                'Scan Attributes
                Dim att As Variant
                att = Entit.GetAttributes()
                'Debug.Print vbTab & (att.Count - 1)
                Dim head_att_no As Long
                'Debug.Print vbTab & "value = " & att(1).EntityName
                For head_att_no = 0 To (att.Count - 1)
                    'Debug.Print vbTab & att(head_att_no).EntityName
                    If att(head_att_no).EntityName = "Attribute" Then
                        'Debug.Print vbTab & vbTab & "Text:   "; att(head_att_no).TextString
                        'Debug.Print vbTab & vbTab & "Tag:    "; att(head_att_no).TagString
                        
                        If att(head_att_no).TagString = "VALUE" Then
                            'Nodes.Add att(head_att_no).TextString
                        End If
                    End If
                Next head_att_no
                
                headIndex = headIndex + 1
                
            'Head IDs
            ElseIf UCase(Entit.layer) = UCase("0b__Head_IDs") Then
                newHeadIDPoint = Entit.insertionPoint
                Debug.Print "Head ID " & Entit.Name & "  " & Entit.Handle
                att = Entit.GetAttributes()
                'Debug.Print vbTab & (att.Count - 1)
                'Dim head_att_no As Long
                'Debug.Print vbTab & "value = " & att(1).EntityName
                For head_att_no = 0 To (att.Count - 1)
                    'Debug.Print vbTab & att(head_att_no).EntityName
                    If att(head_att_no).EntityName = "Attribute" Then
                        'Debug.Print vbTab & vbTab & "Text:   "; att(head_att_no).TextString
                        'Debug.Print vbTab & vbTab & "Tag:    "; att(head_att_no).TagString
                        
                        If att(head_att_no).TagString = "VALUE" Then
                            Set newHeadID = New HeadID
                            newHeadID.Name = att(head_att_no).TextString
                            newHeadID.insertionPoint = newHeadIDPoint 'att(head_att_no).InsertionPoint
                            nodes.Add newHeadID
                        End If
                    End If
                Next head_att_no
            End If

        End If
    
        
    Next Entit
    Debug.Print "Total Pipe Length = " & Format(allPipesLength, "0.0") & " ft  " & headIndex & " heads"
    Debug.Print "All 0.75 inch pipe = " & Format(pipes75Length, "0.0") & " ft"
    Debug.Print "All 1.00 inch pipe = " & Format(pipes1Length, "0.0") & " ft"
    
    
    Debug.Print "Pipe.Count = " & Pipes.Count
    
End Function





Public Function MeasureAllPipe() As Long

    Debug.Print "Layers.Count = " & ActiveDocument.Layers.Count
    Dim lay As IntelliCAD.layer 'AutoCAD.AcadLayer
    Dim myLay As IntelliCAD.layer 'AutoCAD.AcadLayer
    Dim newHeadID As HeadID
    Dim newHead As Head
    Dim newPipe As Pipe
    Dim newHeadIDPoint As Variant
    Dim lngLastCoord As Long
    Dim pipes75Length As Double
    Dim pipes1Length As Double
    
    
    Set nodes = New Collection
    Set Heads = New Collection
    Set Pipes = New Collection
    
    
    Dim allPipesLength As Double, pipeLength As Double, headIndex As Long, riserIndex As Long
    allPipesLength = 0: entitIndex = 1: headIndex = 1: riserIndex = 1
    pipes75Length = 0: pipes1Length = 0
    
    Dim Entit As IntelliCAD.entity 'AutoCAD.AcadEntity
    For Each Entit In ActiveDocument.ModelSpace
        'Pipes
        If Entit.EntityName = "Polyline" Then
            If UCase(Entit.layer) = UCase("0e__Pipes-R-FP") Then
                Set newPipe = New Pipe
                
                newPipe.Diameter = PipeDiameterOfColor(Entit.Color, Entit.layer)
                
                pipeLength = Measure(Entit)
                allPipesLength = allPipesLength + pipeLength
                newPipe.ActualLength = pipeLength
                If newPipe.Diameter = 0.675 Then
                    pipes75Length = pipes75Length + pipeLength
                    'entit.Color = vbYellow
                ElseIf newPipe.Diameter = 0.86 Then
                    pipes1Length = pipes1Length + pipeLength
                    'entit.Color = vbYellow
                Else
                    Debug.Print "This pipe.Diameter is not handled " & newPipe.Diameter
                End If
                
                newPipe.PipeNumber = entitIndex
                lngLastCoord = ((Entit.Coordinates.Count - 1) - 1) / 2
                newPipe.endPoint = Entit.Coordinate(lngLastCoord)
                newPipe.startPoint = Entit.Coordinate(0)
                Pipes.Add newPipe
                Debug.Print "Pipe " & entitIndex & vbTab & ".L=" & Format(pipeLength, "0.00 in") & " " & vbTab & newPipe.Diameter
                entitIndex = entitIndex + 1
            End If
        End If
        
        'Heads & HeadIDs
        If Entit.EntityName = "BlockInsert" Then
            'RFC431616-155
            
            'Fittings, Risers
            If Entit.Name = "PipeFitting" _
                Or Entit.Name = "vpilc" Then
                Set newHead = New Head
                newHead.insertionPoint = Entit.insertionPoint
                newHead.NodeNumber = headIndex
                If Entit.Name = "vpilc" Then
                    newHead.Name = "R." & riserIndex
                    riserIndex = riserIndex + 1
                End If
                Heads.Add newHead
            
            headIndex = headIndex + 1
            
            
            'Heads
            ElseIf UCase(Entit.layer) = UCase("0a__Heads") Then
                Debug.Print "Head " & headIndex & " " & Entit.Name
                
                'If we have a head, add it to the Heads Collection
                '...but not if it's the manifold
                If Entit.Name <> "Manifold_Info" Then
                    Set newHead = New Head
                    newHead.insertionPoint = Entit.insertionPoint
                    newHead.NodeNumber = headIndex
                    newHead.model = Entit.Name
                    Heads.Add newHead
                End If
                
                'Scan Attributes
                Dim att As Variant
                att = Entit.GetAttributes()
                'Debug.Print vbTab & (att.Count - 1)
                Dim head_att_no As Long
                'Debug.Print vbTab & "value = " & att(1).EntityName
                For head_att_no = 0 To (att.Count - 1)
                    'Debug.Print vbTab & att(head_att_no).EntityName
                    If att(head_att_no).EntityName = "Attribute" Then
                        'Debug.Print vbTab & vbTab & "Text:   "; att(head_att_no).TextString
                        'Debug.Print vbTab & vbTab & "Tag:    "; att(head_att_no).TagString
                        
                        If att(head_att_no).TagString = "VALUE" Then
                            'Nodes.Add att(head_att_no).TextString
                        End If
                    End If
                Next head_att_no
                
                headIndex = headIndex + 1
                
            'Head IDs
            ElseIf UCase(Entit.layer) = UCase("0b__Head_IDs") Then
                newHeadIDPoint = Entit.insertionPoint
                Debug.Print "Head ID " & Entit.Name & "  " & Entit.Handle
                att = Entit.GetAttributes()
                'Debug.Print vbTab & (att.Count - 1)
                'Dim head_att_no As Long
                'Debug.Print vbTab & "value = " & att(1).EntityName
                For head_att_no = 0 To (att.Count - 1)
                    'Debug.Print vbTab & att(head_att_no).EntityName
                    If att(head_att_no).EntityName = "Attribute" Then
                        'Debug.Print vbTab & vbTab & "Text:   "; att(head_att_no).TextString
                        'Debug.Print vbTab & vbTab & "Tag:    "; att(head_att_no).TagString
                        
                        If att(head_att_no).TagString = "VALUE" Then
                            Set newHeadID = New HeadID
                            newHeadID.Name = att(head_att_no).TextString
                            newHeadID.insertionPoint = newHeadIDPoint 'att(head_att_no).InsertionPoint
                            nodes.Add newHeadID
                        End If
                    End If
                Next head_att_no
            End If

        End If
    
        
    Next Entit
    Debug.Print "Total Pipe Length = " & Format(allPipesLength / 12, "0.0") & " ft  " & headIndex & " heads"
    Debug.Print "All 0.75 inch pipe = " & Format(pipes75Length / 12, "0.0") & " ft"
    Debug.Print "All 1.00 inch pipe = " & Format(pipes1Length / 12, "0.0") & " ft"
    
    
    Debug.Print "Pipe.Count = " & Pipes.Count
    
End Function








Public Sub ScanEntities() '(OutputFormat As String)

Dim entity As IntelliCAD.entity
Dim Block As IntelliCAD.blockInsert

Dim newNode As Node
Dim newHead As Head
Dim newPipe As Pipe
Dim newTee As Tee
Dim newRiser As Riser
Dim newNodeID As NodeID
Dim newHeadPair As HeadPair
Dim newManifold As Manifold
Dim newFloorTag As FloorTag
Dim newFloorConnector As FloorConnector

Dim pipeLengthTotal As Double
Dim pipeLength As Double

Dim entityIndex As Long
Dim nodeIndex As Long
Dim headIndex As Long
Dim teeIndex As Long
Dim pipeIndex As Long
Dim riserIndex As Long
Dim headPairIndex As Long
Dim manifoldIndex As Long
Dim fixtureIndex As Long

Dim polyline As IntelliCAD.lwPolyline
Dim line As IntelliCAD.line
Dim point As IntelliCAD.point
Dim strHeadModelCode As String

    Set Pipes = New Collection
    Set nodes = New Collection
    Set NodeIDs = New Collection
    Set Heads = New Collection
    Set Tees = New Collection
    Set Risers = New Collection
    Set HeadPairs = New Collection
    Set ManifoldNodes = New Collection
    Set FloorTags = New Collection
    Set FloorConnectors = New Collection
    
    Debug.Print "RunDeleteZeroLengthPipes()..."
    RunDeleteZeroLengthPipes False
    Debug.Print "DeletePipeLabels()..."
    DeletePipeLabels

    manifoldIndex = 0
    headPairIndex = 0
    fixtureIndex = 0
    Dim entityIx As Long
    Dim debugTextE As String
    entityIndex = 0
    For Each entity In ActiveDocument.ModelSpace
        entityIx = entityIx + 1
        Set newNode = New Node
                
       debugTextE = debugTextE & vbCrLf & "Scanning Entity Index: " & entityIx & " " & entity.EntityName & " ID = " & entity.EntityType & " Layer = " & entity.layer & " Color = XXX"
       
        If entityIx Mod 1000 = 0 Then
            Debug.Print debugTextE
            Debug.Print "new block of a total of " & ActiveDocument.ModelSpace.Count & " entities"
            debugTextE = "new block of " & ActiveDocument.ModelSpace.Count & " entities"
            
        End If
        
        'Pipes
        If entity.EntityName = "Polyline" Then
            If UCase(entity.layer) = UCase("0e__Pipes-R-FP") _
                Or UCase(entity.layer) = UCase("0e__Pipes") _
                Or UCase(entity.layer) = UCase("0e__Pipes-R-CW") _
                Or UCase(entity.layer) = UCase("Sprinkler Pipes") _
                Or UCase(Mid(entity.layer, 1, 6)) = UCase("Pipes.") _
                Or UCase(entity.layer) = UCase("Pipes") _
                Or UCase(entity.layer) = UCase("New Pipes") Then
                'Not Old Pipes
                Set polyline = entity
                Set newPipe = New Pipe
                ''If Polyline.color <> Colors.vicGreen Then 'Green is for fixture drops
                    pipeLength = Measure(polyline)
                    newPipe.ActualLength = pipeLength
                    newPipe.Diameter = PipeDiameterOfColor(polyline.Color, entity.layer)
                    newPipe.PipeNumber = pipeIndex
                    newPipe.Name = "P" & pipeIndex
                    
                    
                    newPipe.startPoint = GetStartPoint(polyline)
                    newPipe.endPoint = GetEndPoint(polyline)
                    
                    Set newPipe.polyline = polyline
                    
                    DrawPipeLabel newPipe
                    If Not CollectionContainsName(Pipes, newPipe) Then
                        Pipes.Add newPipe
                        pipeIndex = pipeIndex + 1
                    End If
                ''End If
                
            End If
        
        'Head Pairs
        ElseIf entity.EntityName = "Line" Then
            If UCase(entity.layer) = UCase("0d__2Hd_Calcs") _
                Or UCase(entity.layer) = UCase("TestHeadPair") _
                Or UCase(entity.layer) = UCase("HeadPairs") Then
                Set line = entity
                Set newHeadPair = New HeadPair
                Set newHeadPair.HeadPairLine = line
                
                HeadPairs.Add newHeadPair
                headPairIndex = headPairIndex + 1
            End If
            
        'Manifold Ports
        ElseIf entity.EntityName = "Point" Then
            Set point = entity
            Set newNode = New Node
            newNode.insertionPoint = point.Coordinates
            newNode.NodeNumber = nodeIndex
            newNode.Name = "S.0"
            
            newNode.NodeTypeID = NODE_TYPE_MANIFOLD_PORT
            nodes.Add newNode
            nodeIndex = nodeIndex + 1
            
            ManifoldNodes.Add ConvertToManifold(newNode)
            manifoldIndex = manifoldIndex + 1
                    
        'Heads, Tees, Risers, & NodeIDs ... are all blocks.
        ElseIf entity.EntityName = "BlockInsert" Then    'BlockRef If Block
            Set Block = entity
            
            'Fittings, Risers
            If Block.Name = "PipeFitting" _
                Or Block.Name = "vpilc" _
                Or Block.Name = "PlbgFix" _
                Or Block.Name = "Tee" _
                Or Block.Name = "FloorConnector" Then
                Set newNode = New Node
                Set newTee = New Tee
                
                Set newNode.insertionPoint = Block.insertionPoint
                newNode.NodeNumber = nodeIndex
                
                If Block.Name = "vpilc" Then
                    newNode.Name = "R." & riserIndex & "." & GetAttrib(Block, "VPIPE")
                    newNode.NodeTypeID = NODE_TYPE_RISER
                    riserIndex = riserIndex + 1     'Since these are numbered by this Scanner, we want to make sure they are unique
                
                ElseIf Block.Name = "FloorConnector" Then
                    newNode.Name = "R." & riserIndex
                    newNode.NodeTypeID = NODE_TYPE_RISER
                    riserIndex = riserIndex + 1     'Since these are numbered by this Scanner, we want to make sure they are unique
                    
                ElseIf Block.Name = "PlbgFix" Then
                    Set newNodeID = New NodeID
                    newNodeID.Name = GetAttrib(Block, "SHOWVALUE")
                    newNodeID.insertionPoint = Block.insertionPoint
                    NodeIDs.Add newNodeID

                    newNode.NodeTypeID = NODE_TYPE_FIXTURE
                    newNode.Name = newNodeID.Name
                    
                    fixtureIndex = fixtureIndex + 1
                
                Else
                    newNode.NodeTypeID = NODE_TYPE_TEE
                
                End If
                
                nodes.Add newNode
                nodeIndex = nodeIndex + 1
                
            'New Heads
            ElseIf UCase(entity.layer) = UCase("Heads") Then
                Set newNode = New Node

                Set newNode.insertionPoint = Block.insertionPoint
                
                newNode.NodeNumber = nodeIndex
                strHeadModelCode = GetAttrib(Block, "MODEL")
                newNode.model = strHeadModelCode
                
                newNode.NodeTypeID = NODE_TYPE_HEAD
                nodes.Add newNode
                nodeIndex = nodeIndex + 1
                
                
            'New Nodes, Heads & Tees
            ElseIf UCase(entity.layer) = UCase("0a__Heads") _
                Or UCase(entity.layer) = UCase("Tees") Then
                Set newNode = New Node
                'If we have a head, add it to the NewHeads Collection
                '...but not if it's the manifold
                
                'Heads & Tees:                              They are distinguished from each other by the MatchEntities() function.
                If entity.Name <> "Manifold_Info" Then
                    'This is/should be a head
                    newNode.insertionPoint = entity.insertionPoint
                    newNode.NodeNumber = nodeIndex
                    newNode.model = entity.Name
                    newNode.NodeTypeID = NODE_TYPE_HEAD

                    nodes.Add newNode
                    nodeIndex = nodeIndex + 1
                    
                End If
                
            'Manifold / Source
            ElseIf UCase(entity.layer) = UCase("Manifold Ports") Then
                Set newNode = New Node
                newNode.insertionPoint = entity.insertionPoint
                newNode.NodeNumber = nodeIndex
                newNode.NodeTypeID = NODE_TYPE_MANIFOLD_PORT
                
                nodes.Add newNode
                nodeIndex = nodeIndex + 1
                
                ManifoldNodes.Add ConvertToManifold(newNode)
                manifoldIndex = manifoldIndex + 1
                
            'NodeIDs:                                       These are matched to Heads and Tees, after ScanEntities, by the MatchEntities() function.
            ElseIf UCase(entity.layer) = UCase("0b__Head_IDs") Then
                Set newNodeID = New NodeID
                newNodeID.Name = GetAttrib(entity, "VALUE")
                newNodeID.insertionPoint = entity.insertionPoint
                NodeIDs.Add newNodeID
                
            'HeadLabels (New replacement for nodeid)
            ElseIf UCase(Block.Name) = UCase("HeadLabel") Then
                Set newNodeID = New NodeID
                newNodeID.Name = GetAttrib(entity, "HEADNUMBER")
                Set newNodeID.insertionPoint = Block.insertionPoint
                NodeIDs.Add newNodeID
                If UCase(entity.layer) <> UCase("HeadLabels") Then
                    '2014'ErrorCircle block.InsertionPoint, Colors.vicRed, "Layer not HeadLabels", 1, 0.5, 0.4
                End If
                
            'TeeLabels (New replacement for nodeid)
            ElseIf UCase(Block.Name) = UCase("TeeLabel") Then
                Set newNodeID = New NodeID
                newNodeID.Name = GetAttrib(entity, "TEENUMBER")
                Set newNodeID.insertionPoint = Block.insertionPoint
                NodeIDs.Add newNodeID
                If UCase(entity.layer) <> UCase("TeeLabels") Then
                    '2014'ErrorCircle block.InsertionPoint, Colors.vicRed, "Layer not TeeLabels", 11, 21, 121
                End If
                
            'RiserLabels (New replacement for nodeid)
            ElseIf UCase(Block.Name) = UCase("RiserLabel") Then
                Set newNodeID = New NodeID
                newNodeID.Name = GetAttrib(entity, "RISERNUMBER")
                Set newNodeID.insertionPoint = Block.insertionPoint
                NodeIDs.Add newNodeID
                If UCase(entity.layer) <> UCase("RiserLabels") Then
                    '2014'ErrorCircle block.InsertionPoint, Colors.vicRed, "Layer not RiserLabels", 12, 22, 122
                End If
                
'            'JobData
'            ElseIf block.Name = "JobData" Then
'                Set Job = New Job
'                Job.JobName = GetAttrib(entity, "JOB_NAME")
'                Job.LeadN = GetAttrib(entity, "LEAD_NUMBER")
                
            'Floor Tags
            ElseIf UCase(Block.Name) = UCase("FloorTag") Then
                Set newFloorTag = New FloorTag
                Set newFloorTag.insertionPoint = Block.insertionPoint
                newFloorTag.Name = GetAttrib(entity, "NAME")
                newFloorTag.elevation = GetAttrib(entity, "ELEVATION")
                FloorTags.Add newFloorTag
            
            'Floor Connector
            ElseIf UCase(Block.Name) = UCase("FloorConnector") Then
                Set newFloorConnector = New FloorConnector
                FloorConnectors.Add newFloorConnector
            
            ElseIf UCase(Block.layer) = UCase("0g_Fixtures") Then
                
            
            End If  'BockRef If Block
        End If
        entityIndex = entityIndex + 1
    Next entity
    'End of entity Finding ******************************************************************************************
 
    ScanJobData
 
End Sub
Public Function ScanInterlinks()
Dim Entit As IntelliCAD.entity
Dim entitIndex As Long

    entitIndex = 0
    For Each Entit In ActiveDocument.ModelSpace
        If Entit.EntityName = "BlockInsert" Then
            If UCase(Entit.Name) = UCase("vpilc") Then
            'If entit.Layer = "0e__Pipes" Then
                Debug.Print "Found one"
                'ScanAttrib Entit
                Debug.Print vbTab & GetAttrib(Entit, "VPIPE")
                entitIndex = entitIndex + 1
            End If
        End If
    Next Entit

End Function
Public Function GetAttrib(Entit As IntelliCAD.entity, AttribName As String) As Variant
'Scan Attributes
Dim attribs As IntelliCAD.Attributes
Dim i As Long
Dim Block As IntelliCAD.blockInsert
    
    If TypeOf Entit Is IntelliCAD.blockInsert Then
        Set Block = Entit
        Set attribs = Block.GetAttributes()
        For i = 0 To (attribs.Count - 1)
            If attribs.Item(i).EntityName = "Attribute" Then
                If attribs.Item(i).TagString = AttribName Then
                    GetAttrib = attribs.Item(i).TextString
                End If
            End If
        Next i
    End If

End Function
Public Function SetAttrib(Block As IntelliCAD.blockInsert, AttribName As String, value As String) As Variant
'Public Function SetAttrib(entit As AcadEntity, AttribName As String, Value As String) As Variant
'Scan Attributes
Dim att As IntelliCAD.Attributes
Dim att_no As Long
    
    Set att = Block.GetAttributes()
    For att_no = 0 To (att.Count - 1)
        'If att.Item(att_no).EntityName = "Attribute" Then
            If att(att_no).TagString = AttribName Then
                'GetAttrib = att(att_no).TextString
                att(att_no).TextString = value
                att(att_no).Update
                Exit For
            End If
        'End If
    Next att_no

End Function
Public Function ScanAttrib(Entit As IntelliCAD.entity) As Variant
'Scan Attributes
Dim att As Variant
Dim att_no As Long
    
    att = Entit.GetAttributes()
    'Debug.Print vbTab & (att.Count - 1)
    'Debug.Print vbTab & "value = " & att(1).EntityName
    For att_no = 0 To (att.Count - 1)
        'Debug.Print vbTab & att(head_att_no).EntityName
        If att(att_no).EntityName = "Attribute" Then
            Debug.Print vbTab & vbTab & "Text:   "; att(att_no).TextString
            Debug.Print vbTab & vbTab & "Tag:    "; att(att_no).TagString
            'If att(att_no).TagString = AttribName Then
            '    GetAttrib = att(att_no).TextString
            'End If
        End If
    Next att_no

End Function
Public Function ErrorCircle(centerPoint As point, clr As IntelliCAD.Color, Optional TextString As Variant, Optional Circle1Diameter As Variant, Optional Circle2Diameter As Variant, Optional Circle3Diameter As Variant)
Dim CirCenter As IntelliCAD.point
Dim CirCenterZ As Double
Dim bHasErrorsLayer As Boolean
Dim newCircle As IntelliCAD.Circle
Dim newCircleOuter As IntelliCAD.Circle
Dim newCircleBig As IntelliCAD.Circle
Dim newText As IntelliCAD.text
Dim lay As IntelliCAD.layer

CirCenterZ = 0

    If IsMissing(clr) Then clr = Colors.vicMagenta
    If IsMissing(TextString) Then TextString = ""
    If IsMissing(Circle1Diameter) Then Circle1Diameter = 15
    If IsMissing(Circle2Diameter) Then Circle2Diameter = 17
    If IsMissing(Circle3Diameter) Then Circle3Diameter = 100
    
    CheckAddLayer "SPRINKLER SCANNER ERRORS", Colors.vicCyan
    
    If IsArray(centerPoint) Then
        If VarType(centerPoint(0)) = vbDouble Then
            If (centerPoint.Count - 1) >= 1 Then
                CirCenter.x = centerPoint.x
                CirCenter.y = centerPoint.y
                If (centerPoint.Count - 1) = 1 Then
                    CirCenter.Z = CirCenterZ
                End If
            
                If Circle1Diameter <> 0 Then Set newCircle = ActiveDocument.ModelSpace.AddCircle(CirCenter, Circle1Diameter)
                If Circle2Diameter <> 0 Then Set newCircleOuter = ActiveDocument.ModelSpace.AddCircle(CirCenter, Circle2Diameter)
                If Circle3Diameter <> 0 Then Set newCircleBig = ActiveDocument.ModelSpace.AddCircle(CirCenter, Circle3Diameter)
                If TextString <> "" Then
                    Set newText = ActiveDocument.ModelSpace.AddText(TextString, CirCenter, 10)
                    newText.Rotate CirCenter, DegreesToRadians(30)
                    
                    newText.layer = "SPRINKLER SCANNER ERRORS"
                    
                End If
                                
'                bHasErrorsLayer = False
'                For Each lay In ActiveDocument.Layers
'                    If lay.Name = "SPRINKLER SCANNER ERRORS" Then
'                        bHasErrorsLayer = True
'                        Exit For
'                    End If
'                Next lay
'                If Not bHasErrorsLayer Then ActiveDocument.Layers.Add "SPRINKLER SCANNER ERRORS"
                CheckAddLayer "SPRINKLER SCANNER ERRORS", Colors.vicCyan, False, True
                
                newCircle.Color = clr
                newCircleOuter.Color = clr
                newCircleBig.Color = clr
                newCircle.LineType = "CONTINUOUS"
                newCircleOuter.LineType = "CONTINUOUS"
                newCircleBig.LineType = "CONTINUOUS"
                newCircle.layer = "SPRINKLER SCANNER ERRORS"
                newCircleOuter.layer = "SPRINKLER SCANNER ERRORS"
                newCircleBig.layer = "SPRINKLER SCANNER ERRORS"
                If TextString <> "" Then newText.layer = "SPRINKLER SCANNER ERRORS"
                'newCircle.line
                ErrorCircleCount = ErrorCircleCount + 1
                
            End If
        End If
    End If

End Function


Public Function DeleteErrorCircles()
    Dim lay As IntelliCAD.layer
    Dim bHasErrorsLayer As Boolean
    
    Dim ErrCircle As IntelliCAD.Circle
    Dim ErrText As IntelliCAD.text
    Dim lngCircleCount As Long
    lngCircleCount = 0
    Dim DeleteCircles As New Collection
    Dim DeleteTexts As New Collection
    'Dim lngObjectIndex As Long
    'Dim xObject As AcadObject
    Dim ent As IntelliCAD.entity
    
    
    ErrorCircleCount = 0
    

    
    
    For Each ent In ActiveDocument.ModelSpace
        If ent.EntityName = "Circle" Then
            Set ErrCircle = ent
            If UCase(ErrCircle.layer) = UCase("SPRINKLER SCANNER ERRORS") Then DeleteCircles.Add ErrCircle
        ElseIf ent.EntityName = "Text" Then
            Set ErrText = ent
            If UCase(ErrText.layer) = UCase("SPRINKLER SCANNER ERRORS") Then DeleteTexts.Add ErrText
        End If
    Next ent
    
    lngCircleCount = DeleteCircles.Count
    lngCircleCount = lngCircleCount + DeleteTexts.Count
    
    For Each ErrCircle In DeleteCircles
        ErrCircle.Delete
    Next ErrCircle
    For Each ErrText In DeleteTexts
        ErrText.Delete
    Next ErrText
    
    DeleteErrorCircles = lngCircleCount
End Function

Public Function DeletePipeLabels()
    Dim lay As IntelliCAD.layer
    Dim bHasErrorsLayer As Boolean
    
    Dim labelText As IntelliCAD.text
    Dim ErrText As IntelliCAD.text
    Dim lngLabelCount As Long
    lngLabelCount = 0
    Dim deleteLabels As New Collection

    Dim lngObjectIndex As Long
    'Dim xObject As AcadObject
    Dim entity As IntelliCAD.entity
    
    Dim pipeLabelsLayer  As String
    
    pipeLabelsLayer = "Pipe Labels"
    
    For Each entity In ActiveDocument.ModelSpace
        If entity.EntityName = "Text" Then
            Set labelText = entity
            If UCase(labelText.layer) = UCase(pipeLabelsLayer) Then deleteLabels.Add labelText
            lngLabelCount = lngLabelCount + 1
        End If
    Next entity
    
    
    For Each labelText In deleteLabels
        labelText.Delete
    Next labelText
    
    DeletePipeLabels = lngLabelCount
End Function
Public Function Scan(Section As Long, Optional OutputFormat As Variant) As Long
'Parameters:  Section: 1 Pipes, 9 Nodes, 2 Heads, 3 Tees, 4 Risers, 5 Manifold Nodes

    Dim lay As IntelliCAD.layer
    Dim myLay As IntelliCAD.layer
    Dim newNodeID As NodeID
    Dim newHead As Head
    Dim newPipe As Pipe
    Dim newHeadIDPoint As Variant
    Dim lngLastCoord As Long
    Dim lngSourceNumber As Long
    lngSourceNumber = 0
    
    Dim xPipe As Pipe
    Dim xNodeID As NodeID
    Dim xHead As Head
    Dim newTee As Head
    Dim xNode As Head
    Dim point As IntelliCAD.point
    Dim newHeadPairLine As IntelliCAD.line
    Dim xHeadPairLine As IntelliCAD.line
    Dim newHeadPair As HeadPair
    Dim xHeadPair As HeadPair
    Dim lngHeadPairIndex As Long
    lngHeadPairIndex = 0
    
    Set nodes = New Collection
    Set Heads = New Collection
    Set Pipes = New Collection
    Set Risers = New Collection
    Set Tees = New Collection
    'Set Things = New Collection
    
    
    Dim ManifoldNodes As New Collection
    Dim ManifoldPoints As New Collection
    Dim NewHeads As New Collection
    Dim NodeIDs As New Collection
    Dim HeadPairLines As New Collection
    Dim HeadPairs As New Collection
    Dim BadHeadPairs As New Collection
    
    Dim ResNodes As Collection
    Dim ResPipes As Collection
    Dim newResNode As ResNode
    Dim newResPipe As ResPipe
    
    Dim acc As Double, pipeLength As Double, headIndex As Long, riserIndex As Long
    Dim entitIndex  As Long
    Dim dblErrorElevation As Double
        
    'Res Pipe Variables
    Dim StartNodeName As String
    Dim EndNodeName As String
    Dim StartNodeNameFull As String
    Dim EndNodeNameFull As String
    
    Dim strHeadAName As String
    Dim strHeadBName As String
    
    Dim strPipeMessage As String
    Dim strCalculationName As String
    
    acc = 0: entitIndex = 1: headIndex = 1: riserIndex = 1000
    
    Dim Entit As IntelliCAD.entity
    
    If IsMissing(OutputFormat) Then OutputFormat = "H"
    If OutputFormat = "Hydrauculator" Then OutputFormat = "H"
    If OutputFormat = "Resden" Then OutputFormat = "R"
    If OutputFormat = "Resden44" Then OutputFormat = "R"
    
    
    DeleteErrorCircles
    
    'Begin Entity Finding Loop **********************************************************************
    For Each Entit In ActiveDocument.ModelSpace
        Set newHead = New Head
    
        'Pipes
        If Entit.EntityName = "Polyline" Then
            If UCase(Entit.layer) = UCase("0e__Pipes-R-FP") Then
                pipeLength = Measure(Entit)
                acc = acc + pipeLength
                'Dim xxx As IntelliCAD.Polyline
                'xxx.Coordinates  trace pipe out... This is an array of doubles.
                
                Set newPipe = New Pipe
                newPipe.ActualLength = pipeLength
                newPipe.Diameter = PipeDiameterOfColor(Entit.Color, Entit.layer)
                newPipe.PipeNumber = entitIndex
                newPipe.Name = "P" & entitIndex
                lngLastCoord = ((Entit.Coordinates.Count - 1) - 1) / 2
                newPipe.endPoint = Entit.Coordinate(lngLastCoord)
                newPipe.startPoint = Entit.Coordinate(0)
                If Not CollectionContainsName(Pipes, newPipe) Then
                    Pipes.Add newPipe
                End If
                'Debug.Print "Pipe " & entitIndex & vbTab & ".L=" & Format(pipeLength, "0.00 in") & " " & vbTab & newPipe.Diameter
                entitIndex = entitIndex + 1
            End If
        'Head Pairs
        ElseIf Entit.EntityName = "Line" Then
            If UCase(Entit.layer) = UCase("0d__2Hd_Calcs") Then
                Set newHeadPairLine = Entit
                Set newHeadPair = New HeadPair
                Set newHeadPair.HeadPairLine = newHeadPairLine
                HeadPairs.Add newHeadPair
                lngHeadPairIndex = lngHeadPairIndex + 1
            End If
        End If
        
        'Heads & NodeIDs
        If Entit.EntityName = "BlockInsert" Then
            'RFC431616-155
            Dim b As IntelliCAD.blockInsert
            Set b = Entit
            
            'Fittings, Risers
            If b.Name = "PipeFitting" _
                Or b.Name = "vpilc" Then
                Set newTee = New Head
                newTee.insertionPoint = b.insertionPoint
                newTee.NodeNumber = headIndex
                
                If Entit.Name = "vpilc" Then
                    newTee.Name = "R." & riserIndex & "." & GetAttrib(Entit, "VPIPE")
                    newTee.insertionPoint = b.insertionPoint
                    

                    'Risers.Add newHead
                    'Nodes.Add newHead
                    riserIndex = riserIndex + 1
                End If
                nodes.Add newTee
            
            headIndex = headIndex + 1
            
            
            'New Nodes, Heads & Tees
            ElseIf UCase(Entit.layer) = UCase("0a__Heads") Then
                'Debug.Print "Head " & headIndex & " " & entit.Name

                
                'If we have a head, add it to the NewHeads Collection
                '...but not if it's the manifold
                If UCase(Entit.Name) <> UCase("Manifold_Info") Then
                    'This is/should be a head
                    '''Set newHead = New Head
                    newHead.insertionPoint = Entit.insertionPoint
                    newHead.NodeNumber = headIndex
                    newHead.model = Entit.Name
                    nodes.Add newHead
                Else
                    'TODO: *** Manifold ?
                End If
                
                'Scan Attributes
                ' ?? Getattrib(Entit,"VALUE")
                headIndex = headIndex + 1
                
            'NodeIDs
            ElseIf UCase(Entit.layer) = UCase("0b__Head_IDs") Then
                Set newNodeID = New NodeID
                newNodeID.Name = GetAttrib(Entit, "VALUE")
                newNodeID.insertionPoint = Entit.insertionPoint
                NodeIDs.Add newNodeID
            End If
        End If
        'Manifold Ports
        If UCase(Entit.layer) = UCase("0a__Heads") And Entit.EntityName = "Point" Then
            Set point = Entit
            newHead.insertionPoint = point.Coordinates
            If OutputFormat = "R" Then
                newHead.Name = "S.0"
            Else
                newHead.Name = "S.1"
            End If
            nodes.Add newHead
            ManifoldNodes.Add newHead
            lngSourceNumber = lngSourceNumber + 1
        End If
    Next Entit
    'End of Entity Finding ******************************************************************************************
    
    Debug.Print "//Total Pipe Length = " & Format(acc, "0.0") & " ft  " & headIndex & " heads"
    
    'Debug.Print "Pipe.Count = " & Pipes.Count
    
    'Debug.Print "Nodes.Count = " & Nodes.Count

    
    'For Each xHead In Heads
    '    If Section = 0 Then Debug.Print "junctions.Add( Junction.NewHead(""" & xHead.Name & """, 108, 13, 9.14)); //" & xHead.InsertionPoint.x & ", " & xHead.InsertionPoint.y
    'Next xHead
    
    
    'Begin Entity Matching *********************************************************************************************
    'Debug.Print "Nodes.Count = " & Nodes.Count
    For Each xNode In nodes
        
        'NodeID Label to Node Matcher If Block
        If xNode.Name Like "R.*" Then
            'Risers do not have NodeID labels
            Risers.Add xNode
        'Non Risers: Heads, Tees
        Else
            'Match each Node to a NodeID label
            For Each xNodeID In NodeIDs
                If xNodeID.insertionPoint.y = xNode.insertionPoint.y _
                    And xNodeID.insertionPoint.Z = xNode.insertionPoint.Z _
                    And xNodeID.insertionPoint.x = xNode.insertionPoint.x Then
                    xNode.Name = xNodeID.Name
                    
                    If xNode.Name Like "H.[0-9]*" Then
                        'TODO: newHead.Model = Entit.Name
                        'ScanAttrib Entit
                        If Not CollectionContainsName(Heads, xNode) Then
                            Heads.Add xNode
                        End If
                        
                    ElseIf xNode.Name Like "T.[0-9]*" _
                        Or xNode.Name Like "T[0-9]*" Then
                        If Not CollectionContainsName(Tees, xNode) Then
                            Tees.Add xNode
                        End If
                    End If
                End If
            Next xNodeID
        End If
            
        
        'Match Nodes with pipe ends
        For Each xPipe In Pipes
            If xNode.insertionPoint.x = xPipe.startPoint(0) _
                And xNode.insertionPoint.y = xPipe.startPoint(1) Then
                    'Debug.Print "Pipe Match " & xPipe.PipeNumber & " " & xPipe.ActualLength
                    Set xPipe.StartNode = xNode
            End If
            If xNode.insertionPoint.x = xPipe.endPoint(0) _
                And xNode.insertionPoint.y = xPipe.endPoint(1) Then
                    'Debug.Print "Pipe Match " & xPipe.PipeNumber & " " & xPipe.ActualLength
                    Set xPipe.EndNode = xNode
            End If
        Next xPipe
        
        '''Debug.Print "Node " & xNode.Name & " 100'    //" & Format(xNode.InsertionPoint.x / 12, "0.0") & ", " & Format(xNode.InsertionPoint.y / 12, "0.0") & ", "; Format(xNode.InsertionPoint.z / 12, "0.0")
        If Trim(xNode.Name) = "" Then
            dblErrorElevation = PointElevation(xNode.insertionPoint)
            Debug.Print "Missing Name at " & xNode.insertionPoint.x & ", " & xNode.insertionPoint.y & ", " & xNode.insertionPoint.Z & ", Elevation = " & dblErrorElevation & " make sure the insertion point of the Head or Tee label lines up with the head or tee itself."
            '2014'ErrorCircle xNode.InsertionPoint, Colors.vicCyan, "Missing Name"
        End If
    
    Next xNode
    'End Entity Matching Loop *******************************************************************************************
    
    'Start Match HeadPair to Heads For Loop ******************************************************************************
    For Each xHeadPair In HeadPairs
        For Each xHead In Heads
            If PointsMatch(xHeadPair.HeadPairLine.startPoint, xHead.insertionPoint) Then
                Set xHeadPair.HeadA = xHead
            End If
            If PointsMatch(xHeadPair.HeadPairLine.endPoint, xHead.insertionPoint) Then
                Set xHeadPair.HeadB = xHead
            End If
        Next xHead
    Next xHeadPair
    For Each xHeadPair In HeadPairs
        If xHeadPair.HeadA Is Nothing Then
            BadHeadPairs.Add xHeadPair
            '2014' xHeadPair.HeadPairLine.startPoint, Colors.vicCyan, "HP"
        ElseIf xHeadPair.HeadB Is Nothing Then
            BadHeadPairs.Add xHeadPair
            '2014'ErrorCircle xHeadPair.HeadPairLine.EndPoint, Colors.vicCyan, "HP"
        End If
    Next xHeadPair
    
    If BadHeadPairs.Count > 0 Then
        MsgBox "There are " & BadHeadPairs.Count & " head pairs not linked to heads correctly.  " _
        & "A head pair is the same as a '2Hd Calc Connection'", vbCritical, "BadHeadPairs"
    End If
    
    'End of Match HeadPair to Heads For Loop *************************************************************
    
    'Begin Orphan / Error Finding Loop **********************************************************************************
    'Check both ends of each pipe
    Dim strStart As String, strEnd As String
    For Each xPipe In Pipes
        If xPipe.StartNode Is Nothing Then
            strStart = "X"
            '2014' xPipe.startPoint, Colors.vicMagenta, "P" & xPipe.PipeNumber & " StartNode ERROR"
            '2014'ErrorCircle xPipe.EndPoint, Colors.vicWhite, "P" & xPipe.PipeNumber & " StartNode OK"
        Else
            strStart = xPipe.StartNode.Name
        End If
        If xPipe.EndNode Is Nothing Then
            strEnd = "X"
            '2014'ErrorCircle xPipe.EndPoint, Colors.vicMagenta, "P" & xPipe.PipeNumber & " EndNode ERROR"
            '2014'ErrorCircle xPipe.startPoint, Colors.vicWhite, "P" & xPipe.PipeNumber & " EndNode OK"
        Else
            strEnd = xPipe.EndNode.Name
        End If
        'Debug.Print "Pipe P" & xPipe.PipeNumber & " " & strStart & " " & strEnd & " " & xPipe.RoundedLengthFeet & "' " & xPipe.Diameter & " "
        'Print Head Constructors
    Next xPipe
    'End Orphan / Error Finding Loop **********************************************************************************
    
    
    'Start exporting objects
    If OutputFormat = "RDB" Then
        If Not DeleteNodesPipes() Then
            'MsgBox "Are you sure you want to export the data into Resden Reducer 2000.mdb?"
            Exit Function
        Else
            strCalculationName = InputBox("What would you like to name this calculation?", "Export Sprinkler to Resden Reducer", ActiveDocument.Name & " " & Now())
            If strCalculationName = "" Then
                MsgBox "You didn't select a name, the exporting has been cancelled."
                Exit Function
            End If
            
        End If
    End If
    'Print out lists of nodes
    If Section = 9 Or Section = 0 Then
        Debug.Print "//Nodes:"
        For Each xNode In nodes
            xNode.elevation = PointElevation(xNode.insertionPoint)
            'If Section = 0 Or Section = 2 Then Debug.Print "junctions.Add(Junction.NewHead(""" & xNode.Name & """, " & xNode.Elevation & ", 13, 9.14)); //" & xNode.InsertionPoint.x & ", " & xNode.InsertionPoint.y
        Next xNode
    End If
    If Section = 2 Or Section = 0 Then
        Debug.Print "//Heads:"
        Set Heads = ReSortByName(Heads)
        
        If OutputFormat = "RDB" Then
            Set ResNodes = New Collection
        End If
        
        For Each xHead In Heads
            xHead.elevation = PointElevation(xHead.insertionPoint)  'Get elevation of head
            
            If OutputFormat = "H" Then
                Debug.Print "ss.AddNewHead(""" & xHead.Name & """, " & xHead.elevation & ", """ & xHead.model; """);            //" & xHead.insertionPoint.x & ", " & xHead.insertionPoint.y
            
            ElseIf OutputFormat = "R" Then
                Debug.Print ExtractLong(xHead.Name) & ", " & xHead.elevation & ", " & xHead.insertionPoint.x & ", " & xHead.insertionPoint.y & ", """ & xHead.Name & """"
            
            ElseIf OutputFormat = "RDB" Then
                'The debug line doesn't *need* to be here, but it's handy for debugging...
                Debug.Print ExtractLong(xHead.Name) & ", " & xHead.elevation & ", " & xHead.insertionPoint.x & ", " & xHead.insertionPoint.y & ", """ & xHead.Name & """"
                
                Set newResNode = New ResNode
                newResNode.NodeNo = ExtractLong(xHead.Name)
                newResNode.elevation = xHead.elevation
                newResNode.CoordinateX = xHead.insertionPoint.x
                newResNode.CoordinateY = xHead.insertionPoint.y
                newResNode.NodeName = xHead.Name
                
                ResNodes.Add newResNode, CStr(newResNode.NodeNo)
            Else
                Debug.Print "I'm sorry Dave, I don't recognize the format """ & OutputFormat & """"
            End If
        Next xHead
        
        If OutputFormat = "RDB" Then
            Debug.Print "Copying Head collection to database..."
            basData.AddHeads Heads, strCalculationName, 0
        End If
        
        Debug.Print "//Heads Pairs:"
        'Set HeadPairs = ReSortByName(HeadPairs)
        For Each xHeadPair In HeadPairs
            'Debug.Print "HeadPair " & xHeadPair.HeadA.Name & ", " & xHeadPair.HeadB.Name
            If xHeadPair.HeadA Is Nothing Then
                strHeadAName = "ErrorNothing"
            Else
                strHeadAName = xHeadPair.HeadA.Name
            End If
            If xHeadPair.HeadB Is Nothing Then
                strHeadBName = "ErrorNothing"
            Else
                strHeadBName = xHeadPair.HeadB.Name
            End If
            If OutputFormat = "H" Then
                Debug.Print "ss.AddNewHeadGroup(""" & strHeadAName & """, """ & strHeadBName & """);"
            ElseIf OutputFormat = "R" Then
                Debug.Print "ss.HeadGroups.Add(SprinklerSystem.QuickHeadGrouper(ss.Junctions, """ & strHeadAName & """, """ & strHeadBName & """));"
            Else
                Debug.Print "I'm sorry Dave, I don't recognize the format """ & OutputFormat & """"
            End If
        Next xHeadPair
    End If
    If Section = 3 Or Section = 0 Then
        Debug.Print "//Tees:"
        Set Tees = ReSortByName(Tees)
        
        If OutputFormat = "RDB" Then
            Set ResNodes = New Collection
        End If

        For Each xNode In Tees
            xNode.elevation = PointElevation(xNode.insertionPoint)
            'Debug.Print "Tee " & xNode.Name & " Elevation = " & xNode.Elevation & " // " & xNode.InsertionPoint.x & ", " & xNode.InsertionPoint.y
            If OutputFormat = "H" Then
                'Debug.Print "junctions.Add(Junction.NewJunction(""" & xNode.Name & """, " & xNode.Elevation & "));            //" & xNode.InsertionPoint.x & ", " & xNode.InsertionPoint.y
                Debug.Print "ss.AddNewTee(""" & xNode.Name & """, " & xNode.elevation & ");            //" & xNode.insertionPoint.x & ", " & xNode.insertionPoint.y
            ElseIf OutputFormat = "R" Then
                Debug.Print ExtractLong(xNode.Name) & ", " & xNode.elevation & ", " & xNode.insertionPoint.x & ", " & xNode.insertionPoint.y & ", """ & xNode.Name & """"
            ElseIf OutputFormat = "RDB" Then
                Debug.Print ExtractLong(xNode.Name) & ", " & xNode.elevation & ", " & xNode.insertionPoint.x & ", " & xNode.insertionPoint.y & ", """ & xNode.Name & """"
                
                Set newResNode = New ResNode
                newResNode.NodeNo = ExtractLong(xNode.Name)
                newResNode.elevation = xNode.elevation
                newResNode.CoordinateX = xNode.insertionPoint.x
                newResNode.CoordinateY = xNode.insertionPoint.y
                newResNode.NodeName = xNode.Name
                
                ResNodes.Add newResNode, CStr(newResNode.NodeNo)
                
            Else
                Debug.Print "I'm sorry Dave, I don't recognize the format """ & OutputFormat & """"
            End If
        Next xNode
        If OutputFormat = "RDB" Then
            Debug.Print "Copying Tee collection to database..."
            basData.AddNodes ResNodes, strCalculationName, 0
        End If
        
    End If
    If Section = 4 Or Section = 0 Then      'Export riser nodes
        Debug.Print "//Risers:"
        Set Risers = ReSortByName(Risers)
        
        If OutputFormat = "RDB" Then
            Set ResNodes = New Collection
        End If
        
        For Each xNode In Risers
            xNode.elevation = PointElevation(xNode.insertionPoint)
            'Debug.Print "Riser " & xNode.Name & " Elevation = " & xNode.Elevation & " // " & xNode.InsertionPoint.x & ", " & xNode.InsertionPoint.y
            If OutputFormat = "H" Then
                Debug.Print "ss.AddNewRiser(""" & xNode.Name & """, " & xNode.elevation & ");            //" & xNode.insertionPoint.x & ", " & xNode.insertionPoint.y
            ElseIf OutputFormat = "R" Then
                Debug.Print ExtractLong(xNode.Name) & ", " & xNode.elevation & ", " & xNode.insertionPoint.x & ", " & xNode.insertionPoint.y & ", """ & xNode.Name & """"
            ElseIf OutputFormat = "RDB" Then
                Debug.Print ExtractLong(xNode.Name) & ", " & xNode.elevation & ", " & xNode.insertionPoint.x & ", " & xNode.insertionPoint.y & ", """ & xNode.Name & """"
                
                Set newResNode = New ResNode
                newResNode.NodeNo = ExtractLong(xNode.Name)
                newResNode.elevation = xNode.elevation
                newResNode.CoordinateX = xNode.insertionPoint.x
                newResNode.CoordinateY = xNode.insertionPoint.y
                newResNode.NodeName = xNode.Name
                
                ResNodes.Add newResNode, CStr(newResNode.NodeNo)
                                  
            Else
                Debug.Print "I'm sorry Dave, I don't recognize the format """ & OutputFormat & """"
            End If
        Next xNode
        If OutputFormat = "RDB" Then
            Debug.Print "Copying Riser collection to database..."
            basData.AddNodes ResNodes, strCalculationName, 0
        End If
        
    End If
    If Section = 5 Or Section = 0 Then  'Export mainfolds (one manifold but many manifold ports)
        
        If OutputFormat = "RDB" Then
            Set ResNodes = New Collection
        End If
        Dim strCommentFirstManifold As String
        strCommentFirstManifold = ""
        For Each xNode In ManifoldNodes
            xNode.elevation = PointElevation(xNode.insertionPoint)
            'Debug.Print "Manifold " & xNode.Name & " Elevation = " & xNode.Elevation & " // " & xNode.InsertionPoint.x & ", " & xNode.InsertionPoint.y
            If OutputFormat = "H" Then
                'Debug.Print "//junctions.Add(Junction.NewJunction(""" & xNode.Name & """, " & xNode.Elevation & "));            //" & xNode.InsertionPoint.x & ", " & xNode.InsertionPoint.y
                Debug.Print strCommentFirstManifold & "ss.AddNewManifold(""" & xNode.Name & """, " & xNode.elevation & ");            //" & xNode.insertionPoint.x & ", " & xNode.insertionPoint.y
                strCommentFirstManifold = "//"
            ElseIf OutputFormat = "R" Then
                Debug.Print "//junctions.Add(Junction.NewJunction(""" & xNode.Name & """, " & xNode.elevation & "));            //" & xNode.insertionPoint.x & ", " & xNode.insertionPoint.y
            ElseIf OutputFormat = "RDB" Then
                Debug.Print "//junctions.Add(Junction.NewJunction(""" & xNode.Name & """, " & xNode.elevation & "));            //" & xNode.insertionPoint.x & ", " & xNode.insertionPoint.y
                
                Set newResNode = New ResNode
                newResNode.NodeNo = 0 'The manifold is always zero ... 'ExtractLong(xNode.Name)
                newResNode.elevation = xNode.elevation
                newResNode.CoordinateX = xNode.insertionPoint.x
                newResNode.CoordinateY = xNode.insertionPoint.y
                newResNode.NodeName = xNode.Name
                
                If ResNodes.Count = 0 Then ResNodes.Add newResNode, CStr(newResNode.NodeNo)
                
            Else
                Debug.Print "I'm sorry Dave, I don't recognize the format """ & OutputFormat & """"
            End If
                
            
            'basData.AddNodes ResNodes
        Next xNode
        Debug.Print "//Copying Manifold collection to database..."
    End If
    'Print out a list of pipes
    If Section = 1 Or Section = 0 Then
        Set Pipes = ReSortByName(Pipes)
        
        If OutputFormat = "RDB" Then
            Set ResPipes = New Collection
        End If
        
        For Each xPipe In Pipes
        If OutputFormat = "H" Then
            'If Not (xPipe.StartNode Is Nothing Or xPipe.EndNode Is Nothing) Then
            'Debug.Print "pipes.Add(Pipe.NewPipe(""p" & Format(xPipe.PipeNumber, "00") & """, " & _
            '    Format(xPipe.Diameter, "0.000") & ", """ & xPipe.StartNode.Name & """, """ & xPipe.EndNode.Name & """, " & _
            '    Format(xPipe.ActualLength / 12, "0.000") & ", 150, junctions));"
            'Else
                If xPipe.StartNode Is Nothing And xPipe.EndNode Is Nothing Then
                    StartNodeName = "ErrorNothing"
                    EndNodeName = "ErrorNothing"
                    strPipeMessage = "  //ERROR: Two ends nothing"
                ElseIf xPipe.StartNode Is Nothing Then
                    StartNodeName = "ErrorNothing"
                    EndNodeName = xPipe.EndNode.Name
                    strPipeMessage = "  //ERROR: StartNode nothing"
                ElseIf xPipe.EndNode Is Nothing Then
                    StartNodeName = xPipe.StartNode.Name
                    StartNodeNameFull = "ErrorNothing"
                    EndNodeName = "ErrorNothing"
                    strPipeMessage = "  //ERROR: EndNode nothing"
                Else
                    EndNodeName = xPipe.EndNode.Name
                    StartNodeName = xPipe.StartNode.Name
                    strPipeMessage = " "
                End If
                
'            Debug.Print "pipes.Add(Pipe.NewPipe(""p" & Format(xPipe.PipeNumber, "00") & """, " & _
'                Format(xPipe.Diameter, "0.000") & ", """ & StartNodeName & """, """ & EndNodeName & """, " & _
'                Format(xPipe.ActualLength / 12, "0.000") & ", 150, junctions));  " & strPipeMessage
            Debug.Print "ss.AddNewPipe(""p" & Format(xPipe.PipeNumber, "00") & """, " & _
                Format(xPipe.Diameter, "0.000") & ", """ & StartNodeName & """, """ & EndNodeName & """, " & _
                Format(xPipe.ActualLength / 12, "0.000") & ", 150);  " & strPipeMessage
                    
                
                
            'End If
        ElseIf OutputFormat = "R" Or OutputFormat = "RDB" Then
            If xPipe.StartNode Is Nothing And xPipe.EndNode Is Nothing Then
                StartNodeName = "ErrorNothing"
                StartNodeNameFull = "ErrorNothing"
                EndNodeName = "ErrorNothing"
                EndNodeNameFull = "ErrorNothing"
                strPipeMessage = "  //ERROR: Two ends nothing"
                
            ElseIf xPipe.StartNode Is Nothing Then
                StartNodeName = "ErrorNothing"
                StartNodeNameFull = "ErrorNothing"
                EndNodeName = ExtractLong(xPipe.EndNode.Name)
                EndNodeNameFull = xPipe.EndNode.Name
                strPipeMessage = "  //ERROR: StartNode nothing"
            ElseIf xPipe.EndNode Is Nothing Then
                StartNodeName = ExtractLong(xPipe.StartNode.Name)
                StartNodeNameFull = xPipe.StartNode.Name
                EndNodeName = "ErrorNothing"
                EndNodeNameFull = "ErrorNothing"
                strPipeMessage = "  //ERROR: EndNode nothing"
            Else
                EndNodeName = ExtractLong(xPipe.EndNode.Name)
                EndNodeNameFull = xPipe.EndNode.Name
                StartNodeName = ExtractLong(xPipe.StartNode.Name)
                StartNodeNameFull = xPipe.StartNode.Name
                strPipeMessage = " "
            End If
                
            If StartNodeNameFull = "S.1" Then
                StartNodeNameFull = "S.0"
                StartNodeName = 0
            End If
            If EndNodeNameFull = "S.1" Then
                EndNodeNameFull = "S.0"
                EndNodeName = 0
            End If
            
                
            Debug.Print Format(xPipe.PipeNumber, "00") & ", " & _
                StartNodeName & ", " & _
                EndNodeName & ", " & _
                Format(xPipe.ActualLength / 12, "0.0") & ", " & _
                Format(xPipe.Diameter, "0.000") & ", " & _
                ", ," & _
                " 150, """ & StartNodeNameFull & """, """ & EndNodeNameFull & """"
            
            'R and RDB share most code, expect for this part
            If OutputFormat = "RDB" Then
                Set newResPipe = New ResPipe
                newResPipe.Diameter = xPipe.Diameter
                newResPipe.EndNodeName = EndNodeNameFull
                If EndNodeName = "ErrorNothing" Then
                    MsgBox "Error: EndNodeName is nothing for Pipe " & newResPipe.PipeNo & " (Start Node = " & StartNodeName & ")", vbCritical, "Scan"
                Else
                    newResPipe.EndNodeNo = CLng(EndNodeName)
                End If
                newResPipe.Friction = 150
                newResPipe.Length = xPipe.ActualLength / 12
                newResPipe.PipeNo = xPipe.PipeNumber
                newResPipe.StartNodeName = StartNodeNameFull
                If StartNodeName = "ErrorNothing" Then
                    MsgBox "Error: StartNodeName is nothing for Pipe " & newResPipe.PipeNo & " (Start Node = " & EndNodeName & ")", vbCritical, "Scan"
                Else
                    newResPipe.StartNodeNo = CLng(StartNodeName)
                End If
                
                ResPipes.Add newResPipe, CStr(newResPipe.PipeNo)
            End If

        ElseIf OutputFormat = "Rtest" Then
            Debug.Print Format(xPipe.PipeNumber, "00") & vbTab & _
                ExtractLong(xPipe.StartNode.Name) & vbTab & _
                ExtractLong(xPipe.EndNode.Name) & vbTab & _
                Format(xPipe.ActualLength / 12, "0.0") & vbTab & _
                Format(xPipe.Diameter, "0.000") & vbTab & vbTab & _
                vbTab & vbTab & _
                " 150"
        Else
            Debug.Print "I'm sorry Dave, I don't recognize the format """ & OutputFormat & """"
        End If
        Next xPipe
        
        If OutputFormat = "RDB" Then
            Debug.Print "Adding pipes to database..."
            basData.AddPipes ResPipes, 0
        End If
        
    End If
    

    Debug.Print "//Pipes.Count = " & Pipes.Count
    Debug.Print "//Heads.Count = " & Heads.Count
    Debug.Print "//Tees.Count = " & Tees.Count
    Debug.Print "//Risers.Count = " & Risers.Count
    Debug.Print "//ManifoldNodes.Count = " & ManifoldNodes.Count
    
    If OutputFormat = "RDB" Then
        MsgBox "Sprinker System Scan Complete!  Open the ""Resden Reducer 2000.mdb"" Microsoft Access file to complete the reducing.", vbInformation, "Sprinkler System Scan"
    Else
        MsgBox "Sprinker System Scan Complete!", vbInformation, "Scan"
    End If
    
End Function


Public Sub ScanToResDB()
    NewScan 0, "RDB"
End Sub


Public Sub FindOrphans()
Dim strStart As String
Dim strEnd As String
Dim xPipe As Pipe
    
    'Check both ends of each pipe
    For Each xPipe In Pipes
    
        If xPipe.StartNode Is Nothing Then
            strStart = "X"
            '2014'ErrorCircle xPipe.startPoint, Colors.vicMagenta, "P" & xPipe.PipeNumber & " StartNode = Nothing ERROR"
            '2014'ErrorCircle xPipe.EndPoint, Colors.vicWhite
        Else
            strStart = xPipe.StartNode.Name
        End If
        
        If xPipe.EndNode Is Nothing Then
            strEnd = "X"
            '2014'ErrorCircle xPipe.EndPoint, Colors.vicMagenta, "P" & xPipe.PipeNumber & " EndNode = Nothing ERROR"
            '2014'ErrorCircle xPipe.startPoint, Colors.vicWhite
        Else
            strEnd = xPipe.EndNode.Name
        End If
        
    Next xPipe
    

End Sub
Public Function ExportObjects(Section As Long, OutputFormat As String)
Dim strCalculationName As String

Dim xPipe As Pipe
Dim xNode As Node
Dim xTee As Tee
Dim xRiser As Riser
Dim xHead As Head
Dim xHeadPair As HeadPair
Dim xManifold As Manifold

Dim ResNodes As Collection
Dim ResPipes As Collection

Dim newResNode As ResNode
Dim newResPipe As ResPipe

Dim strHeadAName As String
Dim strHeadBName As String
Dim StartNodeName As String
Dim EndNodeName As String
Dim strPipeMessage  As String
Dim StartNodeNameFull As String
Dim EndNodeNameFull  As String

Dim lngCalculationID As Long

    'Start exporting objects
    
    If OutputFormat = "RDB" Or OutputFormat = "DB" Then
        If Not DeleteNodesPipes() Then
            Exit Function
        Else
            strCalculationName = InputBox("What would you like to name this calculation?", "Export Sprinkler to Resden Reducer", ActiveDocument.Name & " " & Now())
            If strCalculationName = "" Then
                MsgBox "You didn't select a name, the exporting has been cancelled."
                Exit Function
            End If
            lngCalculationID = AddCalculation(strCalculationName)
        End If
    End If
    
    If Section = 99 Then
        Debug.Print "Adding all nodes..."
        basData.AddNodes nodes, strCalculationName, lngCalculationID
    End If
    
    'Print out lists of nodes
    If Section = 9 Or Section = 0 Then
        Debug.Print "//Nodes:"
        For Each xNode In nodes
            xNode.elevation = PointElevation(xNode.insertionPoint)
            'If Section = 0 Or Section = 2 Then Debug.Print "junctions.Add(Junction.NewHead(""" & xNode.Name & """, " & xNode.Elevation & ", 13, 9.14)); //" & xNode.InsertionPoint.x & ", " & xNode.InsertionPoint.y
        Next xNode
    End If
    If Section = 2 Or Section = 0 Then
        Debug.Print "//Heads:"
        Set Heads = ReSortByName(Heads)
        
        If OutputFormat = "RDB" Then
            Set ResNodes = New Collection
        End If
        
        For Each xHead In Heads
            xHead.elevation = PointElevation(xHead.insertionPoint)  'Get elevation of head
            
            If OutputFormat = "H" Then
                Debug.Print "ss.AddNewHead(""" & xHead.Name & """, " & xHead.elevation & ", """ & xHead.model; """);            //" & xHead.insertionPoint.x & ", " & xHead.insertionPoint.y
            
            ElseIf OutputFormat = "R" Then
                Debug.Print ExtractLong(xHead.Name) & ", " & xHead.elevation & ", " & xHead.insertionPoint.x & ", " & xHead.insertionPoint.y & ", """ & xHead.Name & """"
            
            ElseIf OutputFormat = "RDB" Then
                'The debug line doesn't *need* to be here, but it's handy for debugging...
                Debug.Print ExtractLong(xHead.Name) & ", " & xHead.elevation & ", " & xHead.insertionPoint.x & ", " & xHead.insertionPoint.y & ", """ & xHead.Name & """"
                
                Set newResNode = New ResNode
                newResNode.NodeNo = ExtractLong(xHead.Name)
                newResNode.elevation = xHead.elevation
                newResNode.CoordinateX = xHead.insertionPoint.x
                newResNode.CoordinateY = xHead.insertionPoint.y
                newResNode.NodeName = xHead.Name
                
                ResNodes.Add newResNode, CStr(newResNode.NodeNo)
            Else
                Debug.Print "I'm sorry Dave, I don't recognize the format """ & OutputFormat & """"
            End If
        Next xHead
        
        If OutputFormat = "RDB" Then
            Debug.Print "Copying Head collection to database..."
            basData.AddHeads Heads, strCalculationName, lngCalculationID
        End If
        
        Debug.Print "//Heads Pairs:"
        'Set HeadPairs = ReSortByName(HeadPairs)
        For Each xHeadPair In HeadPairs
            'Debug.Print "HeadPair " & xHeadPair.HeadA.Name & ", " & xHeadPair.HeadB.Name
            If xHeadPair.HeadA Is Nothing Then
                strHeadAName = "ErrorNothing"
            Else
                strHeadAName = xHeadPair.HeadA.Name
            End If
            If xHeadPair.HeadB Is Nothing Then
                strHeadBName = "ErrorNothing"
            Else
                strHeadBName = xHeadPair.HeadB.Name
            End If
            If OutputFormat = "H" Then
                Debug.Print "ss.AddNewHeadGroup(""" & strHeadAName & """, """ & strHeadBName & """);"
            ElseIf OutputFormat = "R" Then
                Debug.Print "ss.HeadGroups.Add(SprinklerSystem.QuickHeadGrouper(ss.Junctions, """ & strHeadAName & """, """ & strHeadBName & """));"
            Else
                Debug.Print "I'm sorry Dave, I don't recognize the format """ & OutputFormat & """"
            End If
        Next xHeadPair
    End If
    If Section = 3 Or Section = 0 Then
        Debug.Print "//Tees:"
        Set Tees = ReSortByName(Tees)
        
        If OutputFormat = "RDB" Then
            Set ResNodes = New Collection
        End If

        For Each xTee In Tees
            xTee.elevation = PointElevation(xTee.insertionPoint)
            'Debug.Print "Tee " & xNode.Name & " Elevation = " & xNode.Elevation & " // " & xNode.InsertionPoint.x & ", " & xNode.InsertionPoint.y
            If OutputFormat = "H" Then
                'Debug.Print "junctions.Add(Junction.NewJunction(""" & xNode.Name & """, " & xNode.Elevation & "));            //" & xNode.InsertionPoint.x & ", " & xNode.InsertionPoint.y
                Debug.Print "ss.AddNewTee(""" & xTee.Name & """, " & xTee.elevation & ");            //" & xTee.insertionPoint.x & ", " & xTee.insertionPoint.y
            ElseIf OutputFormat = "R" Then
                Debug.Print ExtractLong(xTee.Name) & ", " & xTee.elevation & ", " & xTee.insertionPoint.x & ", " & xTee.insertionPoint.y & ", """ & xTee.Name & """"
            ElseIf OutputFormat = "RDB" Then
                Debug.Print ExtractLong(xTee.Name) & ", " & xTee.elevation & ", " & xTee.insertionPoint.x & ", " & xTee.insertionPoint.y & ", """ & xTee.Name & """"
                
                Set newResNode = New ResNode
                newResNode.NodeNo = ExtractLong(xTee.Name)
                newResNode.elevation = xTee.elevation
                newResNode.CoordinateX = xTee.insertionPoint.x
                newResNode.CoordinateY = xTee.insertionPoint.y
                newResNode.NodeName = xTee.Name
                
                ResNodes.Add newResNode, CStr(newResNode.NodeNo)
                
            Else
                Debug.Print "I'm sorry Dave, I don't recognize the format """ & OutputFormat & """"
            End If
        Next xTee
        If OutputFormat = "RDB" Then
            Debug.Print "Copying Tee collection to database..."
            basData.AddNodes ResNodes, strCalculationName, lngCalculationID
        End If
        
    End If
    If Section = 4 Or Section = 0 Then      'Export riser nodes
        Debug.Print "//Risers:"
        Set Risers = ReSortByName(Risers)
        
        If OutputFormat = "RDB" Then
            Set ResNodes = New Collection
        End If
        
        For Each xRiser In Risers
            xRiser.elevation = PointElevation(xRiser.insertionPoint)
            'Debug.Print "Riser " & xNode.Name & " Elevation = " & xNode.Elevation & " // " & xNode.InsertionPoint.x & ", " & xNode.InsertionPoint.y
            If OutputFormat = "H" Then
                Debug.Print "ss.AddNewRiser(""" & xRiser.Name & """, " & xRiser.elevation & ");            //" & xRiser.insertionPoint.x & ", " & xRiser.insertionPoint.y
            ElseIf OutputFormat = "R" Then
                Debug.Print ExtractLong(xRiser.Name) & ", " & xRiser.elevation & ", " & xRiser.insertionPoint.x & ", " & xRiser.insertionPoint.y & ", """ & xRiser.Name & """"
            ElseIf OutputFormat = "RDB" Then
                Debug.Print ExtractLong(xRiser.Name) & ", " & xRiser.elevation & ", " & xRiser.insertionPoint.x & ", " & xRiser.insertionPoint.y & ", """ & xRiser.Name & """"
                
                Set newResNode = New ResNode
                newResNode.NodeNo = ExtractLong(xRiser.Name)
                newResNode.elevation = xRiser.elevation
                newResNode.CoordinateX = xRiser.insertionPoint.x
                newResNode.CoordinateY = xRiser.insertionPoint.y
                newResNode.NodeName = xRiser.Name
                
                ResNodes.Add newResNode, CStr(newResNode.NodeNo)
                
            ElseIf OutputFormat = "DB" Then
                Debug.Print ExtractLong(xRiser.Name) & ", " & xRiser.elevation & ", " & xRiser.insertionPoint.x & ", " & xRiser.insertionPoint.y & ", """ & xRiser.Name & """"
                        
            Else
                Debug.Print "I'm sorry Dave, I don't recognize the format """ & OutputFormat & """"
            End If
        Next xRiser
        
        If OutputFormat = "RDB" Then
            Debug.Print "Copying Riser collection to database..."
            basData.AddResNodes ResNodes, strCalculationName, lngCalculationID
        ElseIf OutputFormat = "DB" Then
            Debug.Print "Copying Riser collection to " & MDB_FILE_NAME & " database..."
            ''basData.AddNodes Risers, strCalculationName, lngCalculationID
        End If
        
    End If
    If Section = 5 Or Section = 0 Then  'Export mainfolds (one manifold but many manifold ports)
        
        If OutputFormat = "RDB" Then
            Set ResNodes = New Collection
        End If
        Dim strCommentFirstManifold As String
        strCommentFirstManifold = ""
        For Each xManifold In ManifoldNodes
            xManifold.elevation = PointElevation(xManifold.insertionPoint)
            If OutputFormat = "H" Then
                Debug.Print strCommentFirstManifold & "ss.AddNewManifold(""" & xManifold.Name & """, " & xManifold.elevation & ");            //" & xManifold.insertionPoint.x & ", " & xManifold.insertionPoint.y
                strCommentFirstManifold = "//"
            ElseIf OutputFormat = "R" Then
                Debug.Print "//junctions.Add(Junction.NewJunction(""" & xManifold.Name & """, " & xManifold.elevation & "));            //" & xManifold.insertionPoint.x & ", " & xManifold.insertionPoint.y
            ElseIf OutputFormat = "RDB" Then
                Debug.Print "//junctions.Add(Junction.NewJunction(""" & xManifold.Name & """, " & xManifold.elevation & "));            //" & xManifold.insertionPoint.x & ", " & xManifold.insertionPoint.y
                
                Set newResNode = New ResNode
                newResNode.NodeNo = 0 'The manifold is always zero ... 'ExtractLong(xNode.Name)
                newResNode.elevation = xManifold.elevation
                newResNode.CoordinateX = xManifold.insertionPoint.x
                newResNode.CoordinateY = xManifold.insertionPoint.y
                newResNode.NodeName = xManifold.Name
                
                If ResNodes.Count = 0 Then ResNodes.Add newResNode, CStr(newResNode.NodeNo)
                
            Else
                Debug.Print "I'm sorry Dave, I don't recognize the format """ & OutputFormat & """"
            End If
                
            
            'basData.AddNodes ResNodes
        Next xManifold
        Debug.Print "//Copying Manifold collection to database..."
    End If
    'Print out a list of pipes
    If Section = 1 Or Section = 0 Then
        Set Pipes = ReSortByName(Pipes)
        
        If OutputFormat = "RDB" Then
            Set ResPipes = New Collection
        End If
        
        For Each xPipe In Pipes
            If OutputFormat = "H" Then
                'If Not (xPipe.StartNode Is Nothing Or xPipe.EndNode Is Nothing) Then
                'Debug.Print "pipes.Add(Pipe.NewPipe(""p" & Format(xPipe.PipeNumber, "00") & """, " & _
                '    Format(xPipe.Diameter, "0.000") & ", """ & xPipe.StartNode.Name & """, """ & xPipe.EndNode.Name & """, " & _
                '    Format(xPipe.ActualLength / 12, "0.000") & ", 150, junctions));"
                'Else
                    If xPipe.StartNode Is Nothing And xPipe.EndNode Is Nothing Then
                        StartNodeName = "ErrorNothing"
                        EndNodeName = "ErrorNothing"
                        strPipeMessage = "  //ERROR: Two ends nothing"
                    ElseIf xPipe.StartNode Is Nothing Then
                        StartNodeName = "ErrorNothing"
                        EndNodeName = xPipe.EndNode.Name
                        strPipeMessage = "  //ERROR: StartNode nothing"
                    ElseIf xPipe.EndNode Is Nothing Then
                        StartNodeName = xPipe.StartNode.Name
                        StartNodeNameFull = "ErrorNothing"
                        EndNodeName = "ErrorNothing"
                        strPipeMessage = "  //ERROR: EndNode nothing"
                    Else
                        EndNodeName = xPipe.EndNode.Name
                        StartNodeName = xPipe.StartNode.Name
                        strPipeMessage = " "
                    End If
                    
    '            Debug.Print "pipes.Add(Pipe.NewPipe(""p" & Format(xPipe.PipeNumber, "00") & """, " & _
    '                Format(xPipe.Diameter, "0.000") & ", """ & StartNodeName & """, """ & EndNodeName & """, " & _
    '                Format(xPipe.ActualLength / 12, "0.000") & ", 150, junctions));  " & strPipeMessage
                Debug.Print "ss.AddNewPipe(""p" & Format(xPipe.PipeNumber, "00") & """, " & _
                    Format(xPipe.Diameter, "0.000") & ", """ & StartNodeName & """, """ & EndNodeName & """, " & _
                    Format(xPipe.ActualLength / 12, "0.000") & ", 150);  " & strPipeMessage
                        
                    
                    
                'End If
            ElseIf OutputFormat = "R" Or OutputFormat = "RDB" Then
                If xPipe.StartNode Is Nothing And xPipe.EndNode Is Nothing Then
                    StartNodeName = "ErrorNothing"
                    StartNodeNameFull = "ErrorNothing"
                    EndNodeName = "ErrorNothing"
                    EndNodeNameFull = "ErrorNothing"
                    strPipeMessage = "  //ERROR: Two ends nothing"
                    
                ElseIf xPipe.StartNode Is Nothing Then
                    StartNodeName = "ErrorNothing"
                    StartNodeNameFull = "ErrorNothing"
                    EndNodeName = ExtractLong(xPipe.EndNode.Name)
                    EndNodeNameFull = xPipe.EndNode.Name
                    strPipeMessage = "  //ERROR: StartNode nothing"
                ElseIf xPipe.EndNode Is Nothing Then
                    StartNodeName = ExtractLong(xPipe.StartNode.Name)
                    StartNodeNameFull = xPipe.StartNode.Name
                    EndNodeName = "ErrorNothing"
                    EndNodeNameFull = "ErrorNothing"
                    strPipeMessage = "  //ERROR: EndNode nothing"
                Else
                    EndNodeName = ExtractLong(xPipe.EndNode.Name)
                    EndNodeNameFull = xPipe.EndNode.Name
                    StartNodeName = ExtractLong(xPipe.StartNode.Name)
                    StartNodeNameFull = xPipe.StartNode.Name
                    strPipeMessage = " "
                End If
                    
                If StartNodeNameFull = "S.1" Then
                    StartNodeNameFull = "S.0"
                    StartNodeName = 0
                End If
                If EndNodeNameFull = "S.1" Then
                    EndNodeNameFull = "S.0"
                    EndNodeName = 0
                End If
                
                    
                Debug.Print Format(xPipe.PipeNumber, "00") & ", " & _
                    StartNodeName & ", " & _
                    EndNodeName & ", " & _
                    Format(xPipe.ActualLength / 12, "0.0") & ", " & _
                    Format(xPipe.Diameter, "0.000") & ", " & _
                    ", ," & _
                    " 150, """ & StartNodeNameFull & """, """ & EndNodeNameFull & """"
                
                'R and RDB share most code, expect for this part
                If OutputFormat = "RDB" Then
                    Set newResPipe = New ResPipe
                    newResPipe.Diameter = xPipe.Diameter
                    newResPipe.EndNodeName = EndNodeNameFull
                    If EndNodeName = "ErrorNothing" Then
                        MsgBox "Error: EndNodeName is nothing for Pipe " & newResPipe.PipeNo & " (Start Node = " & StartNodeName & ")", vbCritical, "Scan"
                    Else
                        newResPipe.EndNodeNo = CLng(EndNodeName)
                    End If
                    newResPipe.Friction = 150
                    newResPipe.Length = xPipe.ActualLength / 12
                    newResPipe.PipeNo = xPipe.PipeNumber
                    newResPipe.StartNodeName = StartNodeNameFull
                    If StartNodeName = "ErrorNothing" Then
                        MsgBox "Error: StartNodeName is nothing for Pipe " & newResPipe.PipeNo & " (Start Node = " & EndNodeName & ")", vbCritical, "Scan"
                    Else
                        newResPipe.StartNodeNo = CLng(StartNodeName)
                    End If
                    
                    ResPipes.Add newResPipe, CStr(newResPipe.PipeNo)
                End If
    
            ElseIf OutputFormat = "Rtest" Then
                Debug.Print Format(xPipe.PipeNumber, "00") & vbTab & _
                    ExtractLong(xPipe.StartNode.Name) & vbTab & _
                    ExtractLong(xPipe.EndNode.Name) & vbTab & _
                    Format(xPipe.ActualLength / 12, "0.0") & vbTab & _
                    Format(xPipe.Diameter, "0.000") & vbTab & vbTab & _
                    vbTab & vbTab & _
                    " 150"
            Else
                Debug.Print "I'm sorry Dave, I don't recognize the format """ & OutputFormat & """"
            End If
        Next xPipe
        
        If OutputFormat = "RDB" Then
            Debug.Print "Adding pipes to database..."
            basData.AddPipes ResPipes, strCalculationName, lngCalculationID
        End If
        
   End If
    

End Function
Public Function CheckAddLayer(layerName As String, LayerColor As PipeColors, Optional PrintLayer As Variant, Optional TurnOnLayer As Variant, Optional LineWeight As Variant)
'changed form a sub to a function 2010-12-22 Marc Durham
    Dim lay As IntelliCAD.layer
    Dim bAlreadyHasLayer As Boolean
    Dim newLayer As IntelliCAD.layer
    If IsMissing(PrintLayer) Then PrintLayer = True
    If IsMissing(TurnOnLayer) Then TurnOnLayer = False 'The layer is only turnned on if this value is true, if it's false it's ignored
    If IsMissing(LineWeight) Then LineWeight = vicLnWtByLwDefault  'acLnWt040  '
    
    If VarType(TurnOnLayer) <> vbBoolean Then
        MsgBox "This variable must be true false, or left blank", vbCritical, "CheckAddLayer"
        Exit Function
    End If
    
    For Each lay In ActiveDocument.Layers
        If UCase(lay.Name) = UCase(layerName) Then
            bAlreadyHasLayer = True
            Set CheckAddLayer = lay
            Exit For
        End If
    Next lay
    If Not bAlreadyHasLayer Then
        Set newLayer = ActiveDocument.Layers.Add(layerName)
        newLayer.Color.ColorIndex = LayerColor
        newLayer.Plottable = PrintLayer
        newLayer.LineWeight = LineWeight
        Set CheckAddLayer = newLayer
        Set lay = newLayer
    End If
    If TurnOnLayer Then lay.LayerOn = True
    
    
    
End Function


Public Function DegreesToRadians(Degrees As Double) As Double

    DegreesToRadians = PI * (Degrees / 180#)
    
End Function
Public Sub FindHead() '(HeadName As String)
    Dim xHead As Head
    Dim HeadName As String
    Dim bHeadFound As Boolean
    
    If Heads.Count = 0 Then
        ScanEntities
        MatchEntities
    End If
    
    HeadName = InputBox("Enter head name to find:", "FindHead()", "H.1")
    If HeadName <> "" Then
        For Each xHead In Heads
            If xHead.Name = HeadName Then
                ActiveDocument.Application.ZoomCenter xHead.insertionPoint, 500
                bHeadFound = True
                Exit For
            End If
        Next xHead
        If Not bHeadFound Then MsgBox "Head " & HeadName & " not found!", vbInformation, "FindHead()"
    End If
    
End Sub
Public Sub FindTee()
    Dim xTee As Tee
    Dim TeeName As String
    Dim bTeeFound As Boolean
    
    If Heads.Count = 0 Then
        ScanEntities
        MatchEntities
    End If
    
    TeeName = InputBox("Enter tee name to find:", "FindTee()", "T.1")
    If TeeName <> "" Then
        For Each xTee In Tees
            If xTee.Name = TeeName Then
                ActiveDocument.Application.ZoomCenter xTee.insertionPoint, 500
                bTeeFound = True
                Exit For
            End If
        Next xTee
        If Not bTeeFound Then MsgBox "Tee " & TeeName & " not found!", vbInformation, "FindTee()"
    End If
End Sub

Public Sub DrawHeadPair()
On Error GoTo ErrTrap
Dim HeadAPoint As Variant
Dim HeadBPoint As Variant
Dim PairLine As IntelliCAD.line


    CheckAddLayer "HeadPairs", vicYellow
    'ActiveDocument.Layers("0e__Pipes-R-FP").LayerOn = False
    
DoAgain:
    HeadAPoint = ActiveDocument.Utility.GetPoint(, "Click the first head")
    HeadBPoint = ActiveDocument.Utility.GetPoint(HeadAPoint, "Click the second head")
    
    Set PairLine = ActiveDocument.ModelSpace.AddLine(HeadAPoint, HeadBPoint)
    PairLine.layer = "HeadPairs"
    'PairLine.Thickness = 5
    PairLine.Color = Colors.vicByLayer
    PairLine.LineType = "ByLayer"
    
    GoTo DoAgain
    
ExitProc:
    'ActiveDocument.Layers("0e__Pipes-R-FP").LayerOn = True
    Exit Sub
ErrTrap:
    If Err.Number = -2147352567 Then 'The user pressed escape
        GoTo ExitProc
    ElseIf Err.Number = -2145320928 Then  'user pressed enter
        GoTo ExitProc
    Else
        MsgBox Err.Number & " " & Err.Description, vbCritical, "DrawHeadPair()"
        GoTo ExitProc
        Resume
    End If
    
End Sub
Public Function GetHeadModels() As Collection
Dim xHeadModel As HeadModel
Dim xHeadModelItem As HeadModelItem
    If headModels Is Nothing Then
        basData.LoadHeadModels

'        Set HeadModels = New Collection
'        AddHeadModel HeadModels, "RFC43", 12, 14, 16, 18
'        AddHeadModel HeadModels, "R30", 12, 14, 16
'        AddHeadModel HeadModels, "VK450", 14, 16, 18, 20
        

    End If
    Set GetHeadModels = headModels
End Function

Public Sub AddHeadModel(Models As Collection, ModelNumber As String, ParamArray Coverages() As Variant)
    Dim x As Long
    Dim xCoverage As HeadModelCoverage
    Dim xModel As HeadModel
    
    If Not Models Is Nothing Then
        Set xModel = New HeadModel
        xModel.HeadModelID = Models.Count + 1
        xModel.ModelNumber = ModelNumber
        Models.Add xModel, xModel.ModelNumber
        Set xModel.Coverages = New Collection
        
        For x = 1 To (Coverages.Count - 1)
            Set xCoverage = New HeadModelCoverage
            xCoverage.HeadModelID = xModel.HeadModelID
            xCoverage.Coverage = Coverages(x)
            xModel.Coverages.Add xCoverage
        Next x
    End If
    
End Sub



Public Sub InsertHead()
Dim startPoint As Variant
Dim blockHead As IntelliCAD.blockInsert
    
    frmHeadModels.Show
    
    
End Sub
Public Sub InsertHeadModel(ModelNumber As String, Coverage As String, SlopeName As String, Temperature As String)
On Error GoTo ErrTrap
    Dim startPoint As Variant
    Dim blockHead As IntelliCAD.blockInsert
    Dim strBlockFileName As String
    Dim strSendCommand As String
    strBlockFileName = "Head.dwg"

DoAgain:
    'StartPoint = ActiveDocument.Utility.GetPoint(, "Click where you want the head.")
    'Set blockHead = ActiveDocument.ModelSpace.InsertBlock(StartPoint, strBlockFileName, 1, 1, 1, 0)
    'SetAttrib blockHead, "MODEL", ModelNumber
    'ActiveActiveDocument.SendCommand "(command ""insert"" ""head.dwg"" pause 1 1 0 """ & ModelNumber & """ ) "
    
    CheckAddLayer "Heads", Colors.vicRed, True, True
    
    strSendCommand = "(inserthead """ & ModelNumber & """ """ & Coverage & """ """ & SlopeName & """ """ & Temperature & """ ) "
    ActiveActiveDocument.SendCommand strSendCommand
    ActiveActiveDocument.Activate
    
'    GoTo DoAgain

ExitProc:
    Exit Sub

ErrTrap:
    If Err.Number = -2147352567 Then 'The user pressed escape
        GoTo ExitProc
    ElseIf Err.Number = -2145320928 Then  'user pressed enter
        GoTo ExitProc
    ElseIf Err.Number = -2145386445 Then 'Can't find block file
        MsgBox "Cannot find the block file " & strBlockFileName, vbCritical, "InsertHeadModel()"
        GoTo ExitProc
    Else
        MsgBox Err.Number & " " & Err.Description, vbCritical, "InsertHeadModel()"
        GoTo ExitProc
    End If
End Sub

Public Sub InsertJobData()
On Error GoTo ErrTrap

    If Job Is Nothing Then ScanJobData
    'Else
    If Not Job Is Nothing Then
        If Job.BlockReference Is Nothing Then
            ScanJobData
        ElseIf Job.BlockReference.EntityType = 0 Then
            Set Job = Nothing
        Else
            EditJobData
        End If
    End If
    If Job Is Nothing Then AddJobData

ExitSub:
    Exit Sub
    
ErrTrap:
    If Err.Number = -2145386481 Then
        AddJobData
        Resume ExitSub
    ElseIf Err.Number = -2145386420 Then
        AddJobData
        Resume ExitSub
    Else
        MsgBox Err.Description, vbCritical, "Scanner.InsertJobData()"
        Resume ExitSub
        Resume
    End If
    
End Sub
Public Sub AddJobData()
On Error GoTo ErrTrap
    Dim startPoint As Variant
    Dim blockHead As IntelliCAD.blockInsert
    Dim strBlockFileName As String
    strBlockFileName = "JobData.dwg"
    Dim zeroZero(0 To 2) As Double
    
    zeroZero(0) = 0
    zeroZero(1) = 0
    zeroZero(2) = 0
    
    'startPoint = ActiveDocument.Utility.GetPoint(, "Click where you want the job data")
    startPoint = zeroZero
    
    Set blockHead = ActiveDocument.ModelSpace.InsertBlock(startPoint, strBlockFileName, 1, 1, 1, 0)
    
    frmJobData.Show
    
    
    
'    SetAttrib blockHead, "JOB_NAME", InputBox("Input Job Name", , "Job Name")
'    SetAttrib blockHead, "LEAD_NUMBER", InputBox("Input Lead Number", , "L-0000")
'    SetAttrib blockHead, "SITE_LOCATION", InputBox("Site Location", , "123 Street, City, State Zip")
'    SetAttrib blockHead, "STATIC_PRESSURE", InputBox("Static Pressure (psi)", , "55")
'    SetAttrib blockHead, "RESIDUAL_PRESSURE", InputBox("Residual Pressure (psi)", , "40")
'    SetAttrib blockHead, "AVAILABLE_FLOW", InputBox("Available Flow (gpm)", , "45")
'    SetAttrib blockHead, "METER_ELEVATION", InputBox("Elevation of Meter (ft)", , "100")
'    SetAttrib blockHead, "METER_PIPE_LENGTH", InputBox("Length of pipe to meter (ft)", , "50")
'    SetAttrib blockHead, "METER_PIPE_INTERNAL_DIAMETER", InputBox("Internal diameter of pipe to meter (inches)", , "0.860")
    CheckAddLayer "Job Data", Colors.vicMagenta, False, True
    blockHead.layer = "Job Data"
    
    

ExitProc:
    Exit Sub

ErrTrap:
    If Err.Number = -2147352567 Then 'The user pressed escape
        GoTo ExitProc
    ElseIf Err.Number = -2145320928 Then  'user pressed enter
        GoTo ExitProc
    ElseIf Err.Number = -2145386445 Then 'Can't find block file
        MsgBox "Cannot find the block file " & strBlockFileName, vbCritical, "InsertJobData()"
        GoTo ExitProc
    Else
        MsgBox Err.Number & " " & Err.Description, vbCritical, "InsertJobData()"
        GoTo ExitProc
        Resume
    End If
    
End Sub

Public Function BoxUpperLeft(InputBox As Variant) As Variant
Dim realBox(0 To 7) As Double
Dim upperLeft(0 To 2) As Double
Dim i As Long
Dim minX As Double
Dim maxY As Double

    If (InputBox.Count - 1) = 7 Then 'And LBound(InputBox) = 0 Then ' Progecad: This LBounddoesn't make sense
        If VarType(InputBox(0)) = vbDouble Then
            minX = InputBox(0)
            For i = 0 To 6 Step 2
                If InputBox(i) < minX Then minX = InputBox(i)
            Next i
            maxY = InputBox(1)
            For i = 1 To 7 Step 2
                If InputBox(i) > maxY Then maxY = InputBox(i)
            Next i
            upperLeft(0) = minX
            upperLeft(1) = maxY
            upperLeft(2) = 0
            BoxUpperLeft = upperLeft
        Else
            Err.Raise 1, "BoxUpperLeft", ""
        End If
    Else
        Err.Raise 1, "BoxUpperLeft", ""
    End If
    
    
    
End Function
Public Function BoxCornerMatch(InputBox As Variant, point As Variant) As Boolean
Dim realBox(0 To 7) As Double
Dim upperLeft(0 To 2) As Double
Dim i As Long
Dim j As Long
Dim minX As Double
Dim maxY As Double
Dim foundMatch As Boolean
'
'    If (Point.Count - 1) = 2 And LBound(Point) = 0 Then
'        For i = 0 To 7
'    End If

    If (InputBox.Count - 1) = 7 Then 'And LBound(InputBox) = 0 Then ' Progecad: Removed
        If VarType(InputBox(0)) = vbDouble Then
            For i = 0 To 6 Step 2
                If InputBox(i) = point(0) Then 'match X
                    For j = 1 To 7 Step 2
                        If InputBox(j) = point(1) Then
                            foundMatch = True
                            Exit For
                        End If
                    Next j
                    If foundMatch Then Exit For
                End If
            Next i
        End If
    End If
    
    BoxCornerMatch = foundMatch
    
End Function
Public Sub ScanToXML()
    ScanToXMLOpen False
End Sub
Public Sub ScanToFSDX()
    ScanToFSDXOpen True '2013-02-10' changed 'False
End Sub
Public Sub ScanAndOpen()
    ScanToFSDXOpen True
End Sub
Public Sub ScanToXMLOpen(OpenIt As Boolean)
Dim NodeIDs As Collection
Dim lngCalculationID As Long
Dim strCalculationName As String
Dim strFilePath As String
Dim strDefaultFilePath As String
Dim lngMsgBoxResult As Long

ReScan:
    DeleteErrorCircles
    ScanEntities
    ScanElevationBoxes
    MatchEntities
    HeadPairMatch
    FindOrphans
    
    If Job Is Nothing Then
        lngMsgBoxResult = MsgBox("There is no job data saved, would you like to add it before you export?", vbYesNoCancel, "ScanToXMLOpen()")
        If lngMsgBoxResult = VbMsgBoxResult.vbYes Then
            Scanner.InsertJobData
            GoTo ReScan
        ElseIf lngMsgBoxResult = VbMsgBoxResult.vbNo Then
            strDefaultFilePath = Job.LeadN & " " & Job.JobName
        Else
            Exit Sub 'Cancelled
        End If
    Else
        strDefaultFilePath = "Calculations"
    End If
    If OpenIt Then
        strDefaultFilePath = strDefaultFilePath & ".TEMP.fss"
    Else
        strDefaultFilePath = strDefaultFilePath & ".fss"
    End If

    If Not OpenIt Then
        strFilePath = ShowSave("Save Fire Sprinkler System File As...", strDefaultFilePath, "File Sprinkler System (*.fss)", "*.fss", ActiveDocument.path)
    Else
        strFilePath = strDefaultFilePath
    End If
    If strFilePath = "" Then
        MsgBox "You didn't select a name, the exporting has been cancelled."
        Exit Sub
    End If
    
    basData.SaveXMLFile strFilePath, _
        basData.SaveJobXML(Scanner.Job) & _
        basData.SaveNodesXML(Scanner.nodes) & _
        basData.SavePipesXML(Scanner.Pipes) & _
        basData.SaveHeadPairsXML(Scanner.HeadPairs)
    
    
    
    If OpenIt Then
        'VBA.Shell "c:\loopcalc.exe """ & strFilePath & """", vbNormalFocus
        VBA.Shell strFilePath, vbNormalFocus
    Else
        MsgBox "Sprinker System Scan Complete!" & vbCrLf & "Open the """ & strFilePath & """ from LoopCalc.", vbInformation, "Sprinkler System Scan"
    End If
    
    
End Sub
Public Sub ScanToFSDXOpen(OpenIt As Boolean)
Dim NodeIDs As Collection
Dim lngCalculationID As Long
Dim strCalculationName As String
Dim strFilePath As String
Dim strDefaultFilePath As String
Dim xmlContents As String

    ScanAndMatchEntities
    
    If ErrorCircleCount > 0 Then
        MsgBox "This fire sprinkler system data file cannot be opened or saved.  There are " & ErrorCircleCount & " errors circled."
        Exit Sub
    End If
    
    If Job Is Nothing Then
        Set Job = New Job
        Job.JobName = InputBox("Job name is missing, what job name would you like to use:", "ScanToFSDXOpen", "Job Name")
        Job.LeadN = InputBox("Lead number is missing, what lead number would you like to use:", "ScanToFSDXOpen", "L-0000")
    End If
    
    strDefaultFilePath = Job.LeadN & " " & Job.JobName
    If OpenIt Then
        'strDefaultFilePath = strDefaultFilePath & ".TEMP.fsdx"
        strDefaultFilePath = strDefaultFilePath & ".fsdx"
    Else
        strDefaultFilePath = strDefaultFilePath & ".fsdx"
    End If

    If Not OpenIt Then
        strFilePath = ShowSave("Save Fire Sprinkler Data File As...", strDefaultFilePath, "File Sprinkler Data (*.fsdx)", "*.fsdx", ActiveDocument.path)
    Else
        strFilePath = ShowSave("Save Fire Sprinkler Data File As...", strDefaultFilePath, "File Sprinkler Data (*.fsdx)", "*.fsdx", ActiveDocument.path)
        'strFilePath = strDefaultFilePath
    End If
    If strFilePath = "" Then
        MsgBox "You didn't select a name, the exporting has been cancelled."
        Exit Sub
    End If
    
'    basData.SaveXMLFile strFilePath, _
'        basData.SaveJobXML(Scanner.Job) & _
'        basData.SaveNodesXML(Scanner.Nodes) & _
'        basData.SavePipesXML(Scanner.Pipes) & _
'        basData.SaveHeadPairsXML(Scanner.HeadPairs)

    xmlContents = FSDXJob(Job)
    xmlContents = xmlContents & FSDXNodes(nodes)
    xmlContents = xmlContents & FSDXPipes(Pipes)
    xmlContents = xmlContents & FSDXHeadPairs(HeadPairs)
    xmlContents = FSDXSystem(xmlContents)
    xmlContents = FSDXFile(xmlContents)
        
    basData.SaveFile strFilePath, xmlContents
    
    
    If OpenIt Then
        'VBA.Shell "c:\loopcalc.exe """ & strFilePath & """", vbNormalFocus
        VBA.Shell "CMD /c """ & strFilePath & """, vbNormalFocus"
    Else
        MsgBox "Sprinker System Scan Complete!" & vbCrLf & "Open the """ & strFilePath & """ from LoopCalc.", vbInformation, "Sprinkler System Scan"
    End If
    
    
End Sub

Public Function GetFolder() As String
    GetFolder = ActiveDocument.path
End Function

Public Sub ShowObjectCountOnLayer(layerName As String)

Dim entity As IntelliCAD.entity
Dim polyline As IntelliCAD.lwPolyline
Dim dblResult(0 To 2) As Double
Dim entityIndex As Long

    entityIndex = 0

    For Each entity In ActiveDocument.ModelSpace
        entityIndex = entityIndex + 1
        If entity.layer = layerName Then
        
            
            If entity.EntityName = "Polyline" Then
                Set polyline = entity
                dblResult(0) = polyline.Coordinates(0)
                dblResult(1) = polyline.Coordinates(1)
                dblResult(2) = polyline.Coordinates(2)
                DrawOneLabel dblResult(0), dblResult(1), "O" & entityIndex, 5
            End If
        End If
    
    Next entity
End Sub

Public Sub TestInboxClick()
Dim testPoint As Variant
    testPoint = ActiveDocument.Utility.GetPoint(, "Click where you want the test point.")
    

End Sub

Public Sub TestShell()
VBA.Shell "C:\loopcalc.exe ", vbNormalFocus
End Sub
Public Sub CheckToolbars()
'Application.MenuGroups(1).Toolbars(25).AddToolbarButton 8, "VK450-20","", "^C^C(command ""insert"" ""Head12x12"" pause 1 1 0 ""VK450-20"" )", 0
Dim mg As AcadMenuGroup
Dim loopCalcMenuGroup As AcadMenuGroup
Dim hasLoopCadMenu As Boolean
Dim hasACADMenu As Boolean
Dim hasToolBar As Boolean
Dim tb As AcadToolbar
Dim headsToolBar As AcadToolbar
Dim hasHeadsToolBar As Boolean
Dim i As Long
Dim newToolBarItem As AcadToolbarItem
Dim tbiVK450_18 As AcadToolbarItem
Dim strModelPrefix As String
Dim strCoverageSuffix As String
Dim strTemperature As String
Dim strIconPath As String

    strModelPrefix = "RFC43"

    If Application.MenuGroups.Count > 0 Then
        For Each mg In Application.MenuGroups
            If mg.Name = "LoopCalcMenu" Then
                hasLoopCadMenu = True
                Set loopCalcMenuGroup = mg
                Exit For
            End If
        
        Next mg
    End If
    
    If Not hasLoopCadMenu Then
        If Application.MenuGroups.Count > 0 Then
            For Each mg In Application.MenuGroups
                If mg.Name = "ACAD" Then
                    hasACADMenu = True
                    Set loopCalcMenuGroup = mg
                    loopCalcMenuGroup.SaveAs SUPPORT_FOLDER & "\LoopCAD", acMenuFileSource
                    Exit For
                End If
            Next mg
        End If
        'loopCalcMenuGroup.Name = "LoopCAD"
        'loopCalcMenuGroup.Type = acBaseMenuGroup
        'loopCalcMenuGroup.MenuFileName = "LoopCADMenu"
        loopCalcMenuGroup.SaveAs SUPPORT_FOLDER & "\LoopCADMenuGroup", acMenuFileSource
        'loopCalcMenuGroup
        'Menu.InsertMenuInBar  cool
'        Application.MenuGroups.Load "LoopCADMenuGroup", True 'Load as a base menu...
    End If
        
    If loopCalcMenuGroup.Toolbars.Count > 0 Then
        For Each tb In loopCalcMenuGroup.Toolbars
            If tb.Name = strModelPrefix Then
                Set headsToolBar = tb
                hasHeadsToolBar = True
                Exit For
            End If
        Next tb
    End If
    
    If Not hasHeadsToolBar Then
        Set headsToolBar = loopCalcMenuGroup.Toolbars.Add(strModelPrefix)
        'headsToolBar.Height = 100
        headsToolBar.HelpString = "LoopCAD Heads Toolbar"
        'headsToolBar.LargeButtons = True
        'headsToolBar.TagString = "LoopCAD Heads Toolbar"
    End If
    
    headsToolBar.Visible = True
    headsToolBar.Float 100, 100, 1
    
    Dim hasVK450_18 As Boolean
    If headsToolBar.Count > 0 Then
        
        For i = (headsToolBar.Count - 1) To 0 Step -1
        'For Each tbi In newToolBar
            Set newToolBarItem = headsToolBar.Item(i)
            If newToolBarItem.Name Like strModelPrefix & "-[0-9][0-9]" Then
                'hasVK450_18 = True
                'headsToolBar.Delete
                newToolBarItem.Delete
            End If
        Next i
        'Next tbi
    End If

    ''If Not hasVK450_18 Then
    Dim newCommand As String
        'newCommand = "(defun c:inserthead (model coverage)(while 1(command ""insert"" (strcat ""Head"" coverage ""x"" coverage) pause 1 1 0 (strcat model ""-"" coverage) ))) "
        'debug.Print newCommand
        'ActiveDocument.SendCommand newCommand
        strTemperature = "155"
        strCoverageSuffix = "12"
        i = 1
        GoSub SetNewToolBarItem
        strCoverageSuffix = "14"
        i = 2
        GoSub SetNewToolBarItem
        strCoverageSuffix = "16"
        i = 3
        GoSub SetNewToolBarItem
        strCoverageSuffix = "18"
        i = 4
        GoSub SetNewToolBarItem
        strCoverageSuffix = "20"
        i = 5
        GoSub SetNewToolBarItem


ExitS:
    Exit Sub
    
'        Set tbiVK450_18 = headsToolBar.AddToolbarButton(1, "RFC43-20", "", "(command ""insert"" ""Head20x20"" pause 1 1 0 ""RFC43-20"" ) ")
'        Set tbiVK450_18 = headsToolBar.AddToolbarButton(1, "RFC43-18", "", "(command ""insert"" ""Head18x18"" pause 1 1 0 ""RFC43-18"" ) ")
'        Set tbiVK450_18 = headsToolBar.AddToolbarButton(1, "RFC43-16", "", "(command ""insert"" ""Head16x16"" pause 1 1 0 ""RFC43-16"" ) ")
'        Set tbiVK450_18 = headsToolBar.AddToolbarButton(1, "RFC43-14", "", "(command ""insert"" ""Head14x14"" pause 1 1 0 ""RFC43-14"" ) ")
'        Set tbiVK450_18 = headsToolBar.AddToolbarButton(1, "RFC43-12", "", "(command ""insert"" ""Head12x12"" pause 1 1 0 ""RFC43-12"" ) ")
    ''End If
'    Set newToolBarItem = headsToolBar.AddToolbarButton(i, strModelPrefix & "-" & strCoverageSuffix, "", _
'        Chr(27) & Chr(27) & "(defun *error* (msg) (princ)(princ ""Done"")(princ)(command ""osmode"" osmode)(command ""-layer"" ""off"" ""HeadCoverage"" """" ))" & _
'        "(setq osmode (getvar ""osmode""))(command ""osmode"" 0 )(command ""-layer"" ""on"" ""HeadCoverage"" """" )" & _
'        "(while 1(command ""insert"" ""Head" & strCoverageSuffix & "x" & strCoverageSuffix & """ pause 1 1 0 """ & strModelPrefix & "-" & strCoverageSuffix & """ )) ")
'    Set newToolBarItem = headsToolBar.AddToolbarButton(i, strModelPrefix & "-" & strCoverageSuffix, "", _
'        Chr(27) & Chr(27) & "(defun *error* (msg) (princ)(princ ""Done"")(princ)(command ""osmode"" osmode))(setq osmode (getvar ""osmode""))(command ""osmode"" 0 )(while 1(command ""insert"" ""Head" & strCoverageSuffix & "x" & strCoverageSuffix & """ pause 1 1 0 """ & strModelPrefix & "-" & strCoverageSuffix & """ )) ")
    
SetNewToolBarItem:
    
    Set newToolBarItem = headsToolBar.AddToolbarButton(i, strModelPrefix & "-" & strCoverageSuffix, "", _
        Chr(27) & Chr(27) & "(inserthead """ & strModelPrefix & """ """ & strCoverageSuffix & """ """ & strCoverageSuffix & """ """ & strTemperature & """) ")
    
    
    strIconPath = ICON_FOLDER & "\" & strCoverageSuffix & ".bmp"
    Dim strDirResult As String
    strDirResult = UCase(VBA.Dir(strIconPath, vbNormal))
    If strDirResult = UCase(strCoverageSuffix & ".bmp") Then
        newToolBarItem.SetBitmaps strIconPath, strIconPath
    End If
    
    
    Return

End Sub
Public Sub LoadToolbars()
'Application.MenuGroups(1).Toolbars(25).AddToolbarButton 8, "VK450-20","", "^C^C(command ""insert"" ""Head12x12"" pause 1 1 0 ""VK450-20"" )", 0
Dim mg As AcadMenuGroup
Dim loopCalcMenuGroup As AcadMenuGroup
Dim hasLoopCadMenu As Boolean
Dim hasACADMenu As Boolean
Dim hasToolBar As Boolean
Dim tb As AcadToolbar
Dim headsToolBar As AcadToolbar
Dim hasHeadsToolBar As Boolean
Dim mainToolBar As AcadToolbar
Dim hasMainToolBar As Boolean
Dim i As Long
Dim newToolBarItem As AcadToolbarItem
Dim tbiVK450_18 As AcadToolbarItem
Dim strModelPrefix As String
Dim strCoverageSuffix As String
Dim strTemperature As String
Dim strIconPath As String
Dim strIconPathResult As String
On Error GoTo ErrTrap

    strModelPrefix = "RFC43"
    
    strIconPathResult = VBA.Dir(ICON_FOLDER, vbDirectory) 'error 68 device unavailable
    If strIconPathResult = "" Then
        MsgBox "The folder where the LoopCAD button icons are stored: " & ICON_FOLDER & " could not be accessed!  Cannot continue.", vbCritical, "Load LoopCAD Toolbars"
        Exit Sub
    End If
    

    If Application.MenuGroups.Count > 0 Then
        For Each mg In Application.MenuGroups
            If mg.Name = "LoopCalcMenu" Then
                hasLoopCadMenu = True
                Set loopCalcMenuGroup = mg
                Exit For
            End If
        
        Next mg
    End If
    
    If Not hasLoopCadMenu Then
        If Application.MenuGroups.Count > 0 Then
            For Each mg In Application.MenuGroups
                If mg.Name = "ACAD" Then
                    hasACADMenu = True
                    Set loopCalcMenuGroup = mg
                    loopCalcMenuGroup.SaveAs SUPPORT_FOLDER & "\LoopCAD", acMenuFileSource
                    Exit For
                End If
            Next mg
        End If
        'loopCalcMenuGroup.Name = "LoopCAD"
        'loopCalcMenuGroup.Type = acBaseMenuGroup
        'loopCalcMenuGroup.MenuFileName = "LoopCADMenu"
        loopCalcMenuGroup.SaveAs SUPPORT_FOLDER & "\LoopCADMenuGroup", acMenuFileSource
        'loopCalcMenuGroup
        'Menu.InsertMenuInBar  cool
'        Application.MenuGroups.Load "LoopCADMenuGroup", True 'Load as a base menu...
    End If
    
    
'***********************
'** Main LoopCAD Toolbar
'***********************
    If loopCalcMenuGroup.Toolbars.Count > 0 Then
        For Each tb In loopCalcMenuGroup.Toolbars
            If tb.Name = "LoopCAD" Then
                Set mainToolBar = tb
                hasMainToolBar = True
                Exit For
            End If
        Next tb
    End If
    
    If Not hasMainToolBar Then
        Set mainToolBar = loopCalcMenuGroup.Toolbars.Add("LoopCAD")
        mainToolBar.HelpString = "Main LoopCAD Toolbar"
    End If
    
    mainToolBar.Visible = True
    mainToolBar.Float 100, 100, 1
  
    'Create main tool bar buttons
    If mainToolBar.Count > 0 Then
        For i = (mainToolBar.Count - 1) To 0 Step -1
            Set newToolBarItem = mainToolBar.Item(i)
            'If newToolBarItem.Name Like "LoopCAD" Then
                newToolBarItem.Delete
            'End If
        Next i
    End If
    If mainToolBar.Count = 0 Then
        Dim tbiLoadHeadsBar As AcadToolbarItem
        Dim tbiJobData As AcadToolbarItem
        Dim tbiInsertTee As AcadToolbarItem
        Dim tbiInsertHeads As AcadToolbarItem
        Dim tbiLoadHeadLoader As AcadToolbarItem
        Dim tbiFindNodes As AcadToolbarItem
        Dim tbiFindHeads As AcadToolbarItem
        Dim tbiFindTees As AcadToolbarItem
        Dim tbiDrawElevationBox As AcadToolbarItem
        Dim tbiCreateHeadPair As AcadToolbarItem
        Dim tbiChangeNodeLabel As AcadToolbarItem
        Dim tbiLabelNodes As AcadToolbarItem
        Dim tbiLabelPipes As AcadToolbarItem
        'Dim tbiScanToFSS As AcadToolbarItem
        Dim tbiScanToFSDXOpen As AcadToolbarItem
        Dim tbiScanToFSDX As AcadToolbarItem
        Dim tbiFloorTag As AcadToolbarItem
        Dim tbiFloorConnector As AcadToolbarItem
        Dim tbiScanAndMatchEntities As AcadToolbarItem
        Dim tbiDeleteZeroLengthPipes As AcadToolbarItem
        Dim tbiBreakPipes As AcadToolbarItem
        
        'Load toolbars
        'Make sure the macro has a space at the end (this is like pressing enter)
        Set tbiLoadHeadsBar = mainToolBar.AddToolbarButton(1, "Load heads tool bar", "", Chr(27) & Chr(27) & "-vbarun LoopCAD.Scanner.LoadToolbars ")
        Set tbiJobData = mainToolBar.AddToolbarButton(2, "Insert Job data", "", Chr(27) & Chr(27) & "-vbaRun LoopCAD.Scanner.InsertJobData ")  'No Esc codes here please
        Set tbiInsertTee = mainToolBar.AddToolbarButton(3, "Insert tee", "", Chr(27) & Chr(27) & "(inserttee) ")
        Set tbiInsertHeads = mainToolBar.AddToolbarButton(4, "Insert heads", "", Chr(27) & Chr(27) & "-vbaRun LoopCAD.Scanner.InsertHead ")
        Set tbiLoadHeadLoader = mainToolBar.AddToolbarButton(5, "Load head loader", "", Chr(27) & Chr(27) & "-vbaRun LoopCAD.Scanner.OpenHeadModelBarLoader ")
        Set tbiFindNodes = mainToolBar.AddToolbarButton(6, "Find a node by name/number", "", Chr(27) & Chr(27) & "-vbaRun LoopCAD.Scanner.FindNode ")
        Set tbiFindHeads = mainToolBar.AddToolbarButton(7, "Find a head by name/number", "", Chr(27) & Chr(27) & "-vbaRun LoopCAD.Scanner.FindHead ")
        Set tbiFindTees = mainToolBar.AddToolbarButton(8, "Find a tee by name/number", "", Chr(27) & Chr(27) & "-vbaRun LoopCAD.Scanner.FindTee ")
        Set tbiDrawElevationBox = mainToolBar.AddToolbarButton(9, "Draw elevation box", "", Chr(27) & Chr(27) & "-vbaRun LoopCAD.Scanner.DrawElevationBox ")
        Set tbiCreateHeadPair = mainToolBar.AddToolbarButton(10, "Create head pair connectors", "", Chr(27) & Chr(27) & "-vbaRun LoopCAD.Scanner.DrawHeadPair ")
        Set tbiChangeNodeLabel = mainToolBar.AddToolbarButton(11, "Change existing head number", "", Chr(27) & Chr(27) & "-vbaRun LoopCAD.Scanner.ChangeHead ")
        Set tbiLabelNodes = mainToolBar.AddToolbarButton(11, "Number & label heads", "", Chr(27) & Chr(27) & "-vbaRun LoopCAD.Scanner.LabelNodes ")
        Set tbiLabelPipes = mainToolBar.AddToolbarButton(11, "Number & label pipes", "", Chr(27) & Chr(27) & "-vbaRun LoopCAD.Scanner.LabelPipes ")
        'Set tbiScanToFSS = mainToolBar.AddToolbarButton(12, "Scan to FSS file", "", Chr(27) & Chr(27) & "-vbaRun LoopCAD.Scanner.ScanToXML "& Chr(27) & Chr(27) )
        'Set tbiScanToFSDXOpen = mainToolBar.AddToolbarButton(13, "Scan to FSDX file and open", "", Chr(27) & Chr(27) & "-vbarun LoopCAD.Scanner.ScanAndOpen ")
        Set tbiScanToFSDX = mainToolBar.AddToolbarButton(14, "Scan to FSDX file", "", Chr(27) & Chr(27) & "-vbarun LoopCAD.Scanner.ScanToFSDX ")
        Set tbiFloorTag = mainToolBar.AddToolbarButton(15, "Insert Floor Tag", "", Chr(27) & Chr(27) & "-vbarun LoopCAD.Scanner.InsertFloorTag ")
        Set tbiFloorConnector = mainToolBar.AddToolbarButton(16, "Insert Floor Connector", "", Chr(27) & Chr(27) & "-vbarun LoopCAD.Scanner.InsertFloorConnector ")
        Set tbiScanAndMatchEntities = mainToolBar.AddToolbarButton(17, "Scan Entities", "", Chr(27) & Chr(27) & "-vbarun LoopCAD.Scanner.RunScanAndMatchEntities  ")
        Set tbiDeleteZeroLengthPipes = mainToolBar.AddToolbarButton(18, "Delete Zero Length Pipes", "", Chr(27) & Chr(27) & "-vbarun LoopCAD.Scanner.RunDeleteZeroLengthPipes  ")
        Set tbiBreakPipes = mainToolBar.AddToolbarButton(19, "Break Pipes", "", Chr(27) & Chr(27) & "-vbarun LoopCAD.Scanner.ScanPipesAndBreakAtNodes  ")
        
        
        tbiLoadHeadsBar.SetBitmaps ICON_FOLDER & "\LoadHeadToolBars.bmp", ICON_FOLDER & "\LoadHeadToolBars.bmp"
        tbiJobData.SetBitmaps ICON_FOLDER & "\InsertJobData.bmp", ICON_FOLDER & "\InsertJobData.bmp"
        tbiInsertTee.SetBitmaps ICON_FOLDER & "\InsertTee.bmp", ICON_FOLDER & "\InsertTee.bmp"
        tbiInsertHeads.SetBitmaps ICON_FOLDER & "\InsertHeads.bmp", ICON_FOLDER & "\InsertHeads.bmp"
        tbiLoadHeadLoader.SetBitmaps ICON_FOLDER & "\LoadHeadLoader.bmp", ICON_FOLDER & "\LoadHeadLoader.bmp"
        tbiFindNodes.SetBitmaps ICON_FOLDER & "\FindNode.bmp", ICON_FOLDER & "\FindNode.bmp"
        tbiFindHeads.SetBitmaps ICON_FOLDER & "\FindHeads.bmp", ICON_FOLDER & "\FindHeads.bmp"
        tbiFindTees.SetBitmaps ICON_FOLDER & "\FindTees.bmp", ICON_FOLDER & "\FindTees.bmp"
        tbiDrawElevationBox.SetBitmaps ICON_FOLDER & "\DrawElevationBox.bmp", ICON_FOLDER & "\DrawElevationBox.bmp"
        tbiCreateHeadPair.SetBitmaps ICON_FOLDER & "\CreateHeadPair.bmp", ICON_FOLDER & "\CreateHeadPair.bmp"
        tbiChangeNodeLabel.SetBitmaps ICON_FOLDER & "\ChangeNodeLabel.bmp", ICON_FOLDER & "\ChangeNodeLabel.bmp"
        tbiLabelNodes.SetBitmaps ICON_FOLDER & "\LabelNodes.bmp", ICON_FOLDER & "\LabelNodes.bmp"
        'tbiLabelPipes.SetBitmaps ICON_FOLDER & "\LabelPipes.bmp", ICON_FOLDER & "\LabelPipes.bmp"
        'tbiScanToFSS.SetBitmaps ICON_FOLDER & "\ScanToFSS.bmp", ICON_FOLDER & "\ScanToFSS.bmp"
        'tbiScanToFSDXOpen.SetBitmaps ICON_FOLDER & "\ScanToFSSOpen.bmp", ICON_FOLDER & "\ScanToFSSOpen.bmp"
        tbiScanToFSDX.SetBitmaps ICON_FOLDER & "\ScanToFSDX.bmp", ICON_FOLDER & "\ScanToFSDX.bmp"
        tbiFloorTag.SetBitmaps ICON_FOLDER & "\FloorTag.bmp", ICON_FOLDER & "\FloorTag.bmp"
        tbiFloorConnector.SetBitmaps ICON_FOLDER & "\FloorConnector.bmp", ICON_FOLDER & "\FloorConnector.bmp"
        tbiScanAndMatchEntities.SetBitmaps ICON_FOLDER & "\ScanEntities.bmp", ICON_FOLDER & "\ScanEntities.bmp"
        tbiDeleteZeroLengthPipes.SetBitmaps ICON_FOLDER & "\DeleteZeroLengthPipes.bmp", ICON_FOLDER & "\DeleteZeroLengthPipes.bmp"
        tbiBreakPipes.SetBitmaps ICON_FOLDER & "\BreakPipes.bmp", ICON_FOLDER & "\BreakPipes.bmp"
        
    End If
    
        
        
'******************
'** Head ToolBar
'*******************
    If loopCalcMenuGroup.Toolbars.Count > 0 Then
        For Each tb In loopCalcMenuGroup.Toolbars
            If tb.Name = strModelPrefix Then
                Set headsToolBar = tb
                hasHeadsToolBar = True
                Exit For
            End If
        Next tb
    End If
    
    If Not hasHeadsToolBar Then
        Set headsToolBar = loopCalcMenuGroup.Toolbars.Add(strModelPrefix)
        'headsToolBar.Height = 100
        headsToolBar.HelpString = "LoopCAD Heads Toolbar"
        'headsToolBar.LargeButtons = True
        'headsToolBar.TagString = "LoopCAD Heads Toolbar"
    End If
    
    headsToolBar.Visible = True
    headsToolBar.Float 155, 100, 1
    
    Dim hasVK450_18 As Boolean
    If headsToolBar.Count > 0 Then
        For i = (headsToolBar.Count - 1) To 0 Step -1
            Set newToolBarItem = headsToolBar.Item(i)
            If newToolBarItem.Name Like strModelPrefix & "-[0-9][0-9]" Then
                newToolBarItem.Delete
            End If
        Next i
    End If

    ''If Not hasVK450_18 Then
    Dim newCommand As String
        'newCommand = "(defun c:inserthead (model coverage)(while 1(command ""insert"" (strcat ""Head"" coverage ""x"" coverage) pause 1 1 0 (strcat model ""-"" coverage) ))) "
        'debug.Print newCommand
        'ActiveDocument.SendCommand newCommand
        strTemperature = "155"
        strCoverageSuffix = "12"
        i = 1
        GoSub SetNewToolBarItem
        strCoverageSuffix = "14"
        i = 2
        GoSub SetNewToolBarItem
        strCoverageSuffix = "16"
        i = 3
        GoSub SetNewToolBarItem
        strCoverageSuffix = "18"
        i = 4
        GoSub SetNewToolBarItem
        strCoverageSuffix = "20"
        i = 5
        GoSub SetNewToolBarItem


ExitS:
    Exit Sub
    
'        Set tbiVK450_18 = headsToolBar.AddToolbarButton(1, "RFC43-20", "", "(command ""insert"" ""Head20x20"" pause 1 1 0 ""RFC43-20"" ) ")
'        Set tbiVK450_18 = headsToolBar.AddToolbarButton(1, "RFC43-18", "", "(command ""insert"" ""Head18x18"" pause 1 1 0 ""RFC43-18"" ) ")
'        Set tbiVK450_18 = headsToolBar.AddToolbarButton(1, "RFC43-16", "", "(command ""insert"" ""Head16x16"" pause 1 1 0 ""RFC43-16"" ) ")
'        Set tbiVK450_18 = headsToolBar.AddToolbarButton(1, "RFC43-14", "", "(command ""insert"" ""Head14x14"" pause 1 1 0 ""RFC43-14"" ) ")
'        Set tbiVK450_18 = headsToolBar.AddToolbarButton(1, "RFC43-12", "", "(command ""insert"" ""Head12x12"" pause 1 1 0 ""RFC43-12"" ) ")
    ''End If
'    Set newToolBarItem = headsToolBar.AddToolbarButton(i, strModelPrefix & "-" & strCoverageSuffix, "", _
'        Chr(27) & Chr(27) & "(defun *error* (msg) (princ)(princ ""Done"")(princ)(command ""osmode"" osmode)(command ""-layer"" ""off"" ""HeadCoverage"" """" ))" & _
'        "(setq osmode (getvar ""osmode""))(command ""osmode"" 0 )(command ""-layer"" ""on"" ""HeadCoverage"" """" )" & _
'        "(while 1(command ""insert"" ""Head" & strCoverageSuffix & "x" & strCoverageSuffix & """ pause 1 1 0 """ & strModelPrefix & "-" & strCoverageSuffix & """ )) ")
'    Set newToolBarItem = headsToolBar.AddToolbarButton(i, strModelPrefix & "-" & strCoverageSuffix, "", _
'        Chr(27) & Chr(27) & "(defun *error* (msg) (princ)(princ ""Done"")(princ)(command ""osmode"" osmode))(setq osmode (getvar ""osmode""))(command ""osmode"" 0 )(while 1(command ""insert"" ""Head" & strCoverageSuffix & "x" & strCoverageSuffix & """ pause 1 1 0 """ & strModelPrefix & "-" & strCoverageSuffix & """ )) ")
    
SetNewToolBarItem:
    
'    Set newToolBarItem = headsToolBar.AddToolbarButton(i, strModelPrefix & "-" & strCoverageSuffix, "", _
'        Chr(27) & Chr(27) & "(inserthead """ & strModelPrefix & """ """ & strCoverageSuffix & """ """ & strCoverageSuffix & """ """ & strTemperature & """) ")
    Set newToolBarItem = headsToolBar.AddToolbarButton(i, strModelPrefix & "-" & strCoverageSuffix, "", _
        Chr(27) & Chr(27) & "(insertheadmodelcode """ & strCoverageSuffix & "x" & strCoverageSuffix & """ """ & strModelPrefix & "-" & strCoverageSuffix & """) ")
    
    
    strIconPath = ICON_FOLDER & "\" & strCoverageSuffix & ".bmp"
    Dim strDirResult As String
    strDirResult = UCase(VBA.Dir(strIconPath, vbNormal))
    If strDirResult = UCase(strCoverageSuffix & ".bmp") Then
        newToolBarItem.SetBitmaps strIconPath, strIconPath
    End If
    
    
    Return

ErrTrap:
    If Err.Number = 68 Then
        MsgBox "The device where the folder for the LoopCAD button icons are stored: " & ICON_FOLDER & " could not be accessed!  Cannot continue.", vbCritical, "Load LoopCAD Toolbars"
        Resume ExitS
    Else
        MsgBox "Error " & Err.Number & " " & Err.Description, vbCritical, "Load LoopCAD Toolbars"
        Resume ExitS
    End If
End Sub


Public Sub LoadHeadToolbars(strModelPrefix As String)
Dim mg As AcadMenuGroup
Dim loopCalcMenuGroup As AcadMenuGroup
Dim hasLoopCadMenu As Boolean
Dim hasACADMenu As Boolean
Dim hasToolBar As Boolean
Dim tb As AcadToolbar
Dim headsToolBar As AcadToolbar
Dim hasHeadsToolBar As Boolean
Dim mainToolBar As AcadToolbar
Dim hasMainToolBar As Boolean
Dim i As Long
Dim newToolBarItem As AcadToolbarItem
Dim tbiVK450_18 As AcadToolbarItem
'Dim strModelPrefix As String
Dim strCoverageSuffix As String
Dim strTemperature As String
Dim strIconPath As String
Dim strIconPathResult As String
On Error GoTo ErrTrap


    
    strIconPathResult = VBA.Dir(ICON_FOLDER, vbDirectory) 'error 68 device unavailable
    If strIconPathResult = "" Then
        MsgBox "The folder where the LoopCAD button icons are stored: " & ICON_FOLDER & " could not be accessed!  Cannot continue.", vbCritical, "Load LoopCAD Toolbars"
        Exit Sub
    End If
    

    If Application.MenuGroups.Count > 0 Then
        For Each mg In Application.MenuGroups
            If mg.Name = "LoopCalcMenu" Then
                hasLoopCadMenu = True
                Set loopCalcMenuGroup = mg
                Exit For
            End If
        
        Next mg
    End If
    
    If Not hasLoopCadMenu Then
        If Application.MenuGroups.Count > 0 Then
            For Each mg In Application.MenuGroups
                If mg.Name = "ACAD" Then
                    hasACADMenu = True
                    Set loopCalcMenuGroup = mg
                    loopCalcMenuGroup.SaveAs SUPPORT_FOLDER & "\LoopCAD", acMenuFileSource
                    Exit For
                End If
            Next mg
        End If
        loopCalcMenuGroup.SaveAs SUPPORT_FOLDER & "\LoopCADMenuGroup", acMenuFileSource
    End If
    
        
        
'******************
'** Head ToolBar
'*******************
    If loopCalcMenuGroup.Toolbars.Count > 0 Then
        For Each tb In loopCalcMenuGroup.Toolbars
            If tb.Name = strModelPrefix Then
                Set headsToolBar = tb
                hasHeadsToolBar = True
                Exit For
            End If
        Next tb
    End If
    
    If Not hasHeadsToolBar Then
        Set headsToolBar = loopCalcMenuGroup.Toolbars.Add(strModelPrefix)
        headsToolBar.HelpString = "LoopCAD Heads Toolbar"

    End If
    
    headsToolBar.Visible = True
    headsToolBar.Float 155, 100, 1
    
    Dim hasVK450_18 As Boolean
    If headsToolBar.Count > 0 Then
        For i = (headsToolBar.Count - 1) To 0 Step -1
            Set newToolBarItem = headsToolBar.Item(i)
            If newToolBarItem.Name Like strModelPrefix & "-[0-9][0-9]" Then
                newToolBarItem.Delete
            End If
        Next i
    End If

    ''If Not hasVK450_18 Then
    Dim newCommand As String
        'newCommand = "(defun c:inserthead (model coverage)(while 1(command ""insert"" (strcat ""Head"" coverage ""x"" coverage) pause 1 1 0 (strcat model ""-"" coverage) ))) "
        'debug.Print newCommand
        'ActiveDocument.SendCommand newCommand
        strTemperature = "155"
        strCoverageSuffix = "12"
        i = 1
        GoSub SetNewToolBarItem
        strCoverageSuffix = "14"
        i = 2
        GoSub SetNewToolBarItem
        strCoverageSuffix = "16"
        i = 3
        GoSub SetNewToolBarItem
        strCoverageSuffix = "18"
        i = 4
        GoSub SetNewToolBarItem
        strCoverageSuffix = "20"
        i = 5
        GoSub SetNewToolBarItem


ExitS:
    Exit Sub
    
SetNewToolBarItem:
    
    Set newToolBarItem = headsToolBar.AddToolbarButton(i, strModelPrefix & "-" & strCoverageSuffix, "", _
        Chr(27) & Chr(27) & "(inserthead """ & strModelPrefix & """ """ & strCoverageSuffix & """ """ & strCoverageSuffix & """ """ & strTemperature & """) ")
    
    
    strIconPath = ICON_FOLDER & "\" & strCoverageSuffix & ".bmp"
    Dim strDirResult As String
    strDirResult = UCase(VBA.Dir(strIconPath, vbNormal))
    If strDirResult = UCase(strCoverageSuffix & ".bmp") Then
        newToolBarItem.SetBitmaps strIconPath, strIconPath
    End If
    
    
    Return

ErrTrap:
    If Err.Number = 68 Then
        MsgBox "The device where the folder for the LoopCAD button icons are stored: " & ICON_FOLDER & " could not be accessed!  Cannot continue.", vbCritical, "Load LoopCAD Toolbars"
        Resume ExitS
    Else
        MsgBox "Error " & Err.Number & " " & Err.Description, vbCritical, "Load LoopCAD Toolbars"
        Resume ExitS
    End If
End Sub
Public Function InsertFloorConnectorTest()
On Error GoTo ErrTrap
    Dim startPoint As Variant
    Dim blockHead As IntelliCAD.blockInsert
    Dim strBlockFileName As String
    Dim strSendCommand As String
    Dim xFloorTag As FloorTag
    
    strBlockFileName = "FloorConnector.dwg"

DoAgain:
    startPoint = ActiveDocument.Utility.GetPoint(, "Click where you want the floor connector.")
    'Set blockHead = ActiveDocument.ModelSpace.InsertBlock(StartPoint, strBlockFileName, 1, 1, 1, 0)
    
    ScanEntities
    
    For Each xFloorTag In FloorTags

        Debug.Print "FloorTags " & xFloorTag.Name

    Next xFloorTag

ExitProc:
    Exit Function

ErrTrap:
    If Err.Number = -2147352567 Then 'The user pressed escape
        GoTo ExitProc
    ElseIf Err.Number = -2145320928 Then  'user pressed enter
        GoTo ExitProc
    ElseIf Err.Number = -2145386445 Then 'Can't find block file
        MsgBox "Cannot find the block file " & strBlockFileName, vbCritical, "InsertHeadModel()"
        GoTo ExitProc
    Else
        MsgBox Err.Number & " " & Err.Description, vbCritical, "InsertHeadModel()"
        GoTo ExitProc
    End If


End Function

Public Function MeasureBetween(startPoint As Variant, endPoint As Variant) As Double

    
    If ((startPoint.Count - 1) = 1 Or (startPoint.Count - 1) = 2) And ((endPoint.Count - 1) = 1 Or (endPoint.Count - 1) = 2) Then
        MeasureBetween = Math.Sqr((startPoint(0) - endPoint(0)) ^ 2 + (startPoint(1) - endPoint(1)) ^ 2)
    Else
        MeasureBetween = 0
    End If
    

End Function
Public Sub TestMeasureBetween()
    Dim s(0 To 2) As Double
    Dim e(0 To 2) As Double
    
    s(0) = 5
    s(1) = 5
    s(2) = 0
    
    e(0) = 10
    e(1) = 10
    e(2) = 0
    
    If Format(MeasureBetween(s, e), "0.0000") = "7.0711" Then
        Debug.Print "MeasureBetween 1 PASS  = " & MeasureBetween(s, e)
    ElseIf Format(MeasureBetween(s, e), "0.0000") = "7.0710" Then
        Debug.Print "MeasureBetween 1 FAIL = " & MeasureBetween(s, e)
    Else
        Debug.Print "MeasureBetween 1 FAIL = " & MeasureBetween(s, e)
    End If
    
    
    s(0) = -5
    s(1) = -5
    s(2) = 0
    
    e(0) = -10
    e(1) = -10
    e(2) = 0
    
    If Format(MeasureBetween(s, e), "0.0000") = "7.0711" Then
        Debug.Print "MeasureBetween 2 PASS  = " & MeasureBetween(s, e)
    ElseIf Format(MeasureBetween(s, e), "0.0000") = "7.0710" Then
        Debug.Print "MeasureBetween 2 FAIL = " & MeasureBetween(s, e)
    Else
        Debug.Print "MeasureBetween 2 FAIL = " & MeasureBetween(s, e)
    End If
    
    
    s(0) = -5
    s(1) = -5
    s(2) = 0
    
    e(0) = 10
    e(1) = -5
    e(2) = 0
    
    If Format(MeasureBetween(s, e), "0.0000") = "15.0000" Then
        Debug.Print "MeasureBetween 2 PASS  = " & MeasureBetween(s, e)
    ElseIf Format(MeasureBetween(s, e), "0.0000") = "7.0710" Then
        Debug.Print "MeasureBetween 2 FAIL = " & MeasureBetween(s, e)
    Else
        Debug.Print "MeasureBetween 2 FAIL = " & MeasureBetween(s, e)
    End If
    
    
End Sub
Public Sub TestFloorTags_All() '(TestPoint As Variant)
    Dim dist As Double
    Dim testPoint As Variant
    Dim xFloorTag As FloorTag
    Dim closestFloorTag As FloorTag
    Dim closestDistance As Double
    
    testPoint = ActiveDocument.Utility.GetPoint(, "Click where you want the test point.")
    
    If FloorTags Is Nothing Then ScanEntities
    If FloorTags.Count <= 0 Then ScanEntities
    
    Debug.Print "FloorTag.Count = " & FloorTags.Count

    If (testPoint.Count - 1) <> 2 Then
        Debug.Print "not a valid test point"
        Exit Sub
    End If
    
    For Each xFloorTag In FloorTags
        dist = MeasureBetween(xFloorTag.insertionPoint, testPoint)
        If closestDistance = 0 Or dist < closestDistance Then
            closestDistance = dist
            Set closestFloorTag = xFloorTag
        End If
        'closestFloorTag
        Debug.Print "FloorTag = " & xFloorTag.Name & " " & xFloorTag.insertionPoint.x & ", " & xFloorTag.insertionPoint.y & ", " & xFloorTag.insertionPoint.Z & " Distance = " & dist
    Next xFloorTag
    
    Debug.Print "The closest point is " & closestFloorTag.Name
    Debug.Print "The closest point to 2 is " & NearestFloorTag(testPoint).Name
    Dim relative As Variant
    
    relative = RelativeCoordinates(closestFloorTag.insertionPoint, testPoint)
    Debug.Print "The relative point is " & relative(0) & ", " & relative(1)
    
    Dim absolute As Variant
    absolute = AddCoordinates(closestFloorTag.insertionPoint, relative)
    Debug.Print "The absolute point is " & absolute(0) & ", " & absolute(1)

    Debug.Print "The clicked point is " & testPoint(0) & ", " & testPoint(1)
    
    Debug.Print "FloorTags.Count = " & FloorTags.Count
    For Each xFloorTag In FloorTags
        'If xFloorTag.Name <> closestFloorTag.Name Then
            relative = RelativeCoordinates(closestFloorTag.insertionPoint, testPoint)
            absolute = AddCoordinates(xFloorTag.insertionPoint, relative)
            ActiveDocument.ModelSpace.InsertBlock absolute, "FloorConnector.dwg", 1, 1, 1, 0
            
            Debug.Print " inserting at " & absolute(0) & ", " & absolute(1) & " for " & xFloorTag.Name
        'End If
    Next xFloorTag
    
    'I need one floor tag per elevation box...
    'From there I can figure out which goes where
End Sub
Public Sub TestAddCoordinates()
    Dim Result As Variant '(0 To 2) As Double
    Dim s(0 To 2) As Double
    Dim e(0 To 2) As Double
    
    s(0) = 10
    s(1) = 10
    s(2) = 0
    
    e(0) = 30
    e(1) = 30
    e(2) = 0
    
    Result = AddCoordinates(s, e)
    
    If Result(0) = 40 And Result(1) = 40 Then
        Debug.Print "TestAddCoordinates 1 PASS"
    Else
        Debug.Print "TestAddCoordinates 1 FAIL"
    End If

    s(0) = -10
    s(1) = -10
    s(2) = 0
    
    e(0) = -30
    e(1) = -30
    e(2) = 0
    
    Result = AddCoordinates(s, e)
    
    If Result(0) = -40 And Result(1) = -40 Then
        Debug.Print "TestAddCoordinates 2 PASS"
    Else
        Debug.Print "TestAddCoordinates 2 FAIL"
    End If

    s(0) = 10
    s(1) = 10
    s(2) = 0
    
    e(0) = -30
    e(1) = -30
    e(2) = 0
    
    Result = AddCoordinates(s, e)
    
    If Result(0) = -20 And Result(1) = -20 Then
        Debug.Print "TestAddCoordinates 3 PASS"
    Else
        Debug.Print "TestAddCoordinates 3 FAIL"
    End If

End Sub
Public Sub TestRelativeCoordinates()
    Dim Result As Variant '(0 To 2) As Double
    Dim s(0 To 2) As Double
    Dim e(0 To 2) As Double
    
    s(0) = 10
    s(1) = 10
    s(2) = 0
    
    e(0) = 30
    e(1) = 30
    e(2) = 0
    
    Result = RelativeCoordinates(s, e)
    
    If Result(0) = 20 And Result(1) = 20 Then
        Debug.Print "TestRelativeCoordinates 1 PASS"
    Else
        Debug.Print "TestRelativeCoordinates 1 FAIL"
    End If
    
    s(0) = -10
    s(1) = -10
    s(2) = 0
    
    e(0) = -30
    e(1) = -30
    e(2) = 0
    
    Result = RelativeCoordinates(s, e)
    
    If Result(0) = -20 And Result(1) = -20 Then
        Debug.Print "TestRelativeCoordinates 2 PASS"
    Else
        Debug.Print "TestRelativeCoordinates 2 FAIL"
    End If
    
    
    s(0) = 10
    s(1) = -10
    s(2) = 0
    
    e(0) = 30
    e(1) = -30
    e(2) = 0
    
    Result = RelativeCoordinates(s, e)
    
    If Result(0) = 20 And Result(1) = -20 Then
        Debug.Print "TestRelativeCoordinates 3 PASS"
    Else
        Debug.Print "TestRelativeCoordinates 3 FAIL"
    End If

    s(0) = -10
    s(1) = 10
    s(2) = 0
    
    e(0) = 30
    e(1) = -30
    e(2) = 0
    
    Result = RelativeCoordinates(s, e)
    
    If Result(0) = 40 And Result(1) = -40 Then
        Debug.Print "TestRelativeCoordinates 4 PASS"
    Else
        Debug.Print "TestRelativeCoordinates 4 FAIL"
    End If

End Sub
Public Function RelativeCoordinates(centerPoint As Variant, testPoint As Variant) As Variant
    Dim Result(0 To 2) As Double
    
    Result(0) = testPoint(0) - centerPoint(0)
    Result(1) = testPoint(1) - centerPoint(1)
    Result(2) = testPoint(2) - centerPoint(2)
    
    RelativeCoordinates = Result
    
End Function
Public Function NearestFloorTag(testPoint As Variant) As FloorTag
    Dim dist As Double
    'Dim TestPoint As Variant
    Dim xFloorTag As FloorTag
    Dim closestFloorTag As FloorTag
    Dim closestDistance As Double
    
    'TestPoint = ActiveDocument.Utility.GetPoint(, "Click where you want the test point.")
    
    If FloorTags Is Nothing Then ScanEntities
    If FloorTags.Count <= 0 Then ScanEntities
    
    If (testPoint.Count - 1) <> 2 Then
        Debug.Print "Not a valid test point!"
        Set NearestFloorTag = Nothing
        Exit Function
    End If
    
    For Each xFloorTag In FloorTags
        dist = MeasureBetween(xFloorTag.insertionPoint, testPoint)
        If closestDistance = 0 Or dist < closestDistance Then
            closestDistance = dist
            Set closestFloorTag = xFloorTag
        End If
        'closestFloorTag
        Debug.Print "FloorTag = " & xFloorTag.Name & " " & xFloorTag.insertionPoint.x & ", " & xFloorTag.insertionPoint.y & ", " & xFloorTag.insertionPoint.Z & " Distance = " & dist
    Next xFloorTag
    
    ' Debug.Print "The closest point is " & closestFloorTag.Name
    Set NearestFloorTag = closestFloorTag
    
End Function
Public Function AddCoordinates(centerPoint As Variant, testPoint As Variant) As Variant
    Dim Result(0 To 2) As Double
    
    Result(0) = testPoint(0) + centerPoint(0)
    Result(1) = testPoint(1) + centerPoint(1)
    Result(2) = testPoint(2) + centerPoint(2)
    
    AddCoordinates = Result
    
End Function
Public Sub InsertFloorTag()
On Error GoTo ErrTrap
    Dim startPoint As Variant
    Dim blockHead As IntelliCAD.blockInsert
    Dim strBlockFileName As String
    'Dim strSendCommand As String
    Dim strFloorName As String
    Dim strElevation As String
    
    strBlockFileName = "FloorTag.dwg"

DoAgain:
    startPoint = ActiveDocument.Utility.GetPoint(, "Click where you want the floor tag.")
    strFloorName = InputBox("Please input the floor name:", "Floor Tag Floor Name", "Main Floor")
    If strFloorName = "" Then Exit Sub
    strElevation = InputBox("Please input the elevation for this floor (sprinkler head level):", "Floor Tag Elevation", "108")
    If strElevation = "" Then Exit Sub
    
    Set blockHead = ActiveDocument.ModelSpace.InsertBlock(startPoint, strBlockFileName, 1, 1, 1, 0)
    SetAttrib blockHead, "NAME", strFloorName
    SetAttrib blockHead, "ELEVATION", strElevation
    CheckAddLayer "Floor Tags", Colors.vicCyan, False, True
    blockHead.layer = "Floor Tags"
ExitProc:
    Exit Sub

ErrTrap:
    If Err.Number = -2147352567 Then 'The user pressed escape
        GoTo ExitProc
    ElseIf Err.Number = -2145320928 Then  'user pressed enter
        GoTo ExitProc
    ElseIf Err.Number = -2145386445 Then 'Can't find block file
        MsgBox "Cannot find the block file " & strBlockFileName, vbCritical, "InsertFloorTag()"
        GoTo ExitProc
    Else
        MsgBox Err.Number & " " & Err.Description, vbCritical, "InsertFloorTag()"
        GoTo ExitProc
        Resume
    End If


End Sub

Public Sub InsertFloorConnector()
    Dim dist As Double
    Dim xFloorTag As FloorTag
    Dim closestFloorTag As FloorTag
    Dim closestDistance As Double
    Dim eBox As ElevationBox
    Dim Block As IntelliCAD.blockInsert
    Dim testPoint As IntelliCAD.point

    Set testPoint = ActiveDocument.Utility.GetPoint(, "Click where you want the test point.")
    
    If FloorTags Is Nothing Then ScanEntities
    If FloorTags.Count <= 0 Then ScanEntities
    If ElevationBoxes Is Nothing Then ScanElevationBoxes
    If ElevationBoxes.Count <= 0 Then ScanElevationBoxes
    
    
    Debug.Print "FloorTag.Count = " & FloorTags.Count

    If IsNull(testPoint) Then
        Debug.Print "not a valid test point"
        Exit Sub
    End If
    
    Set closestFloorTag = GetElevationBoxFloorTag(testPoint)
    
    If closestFloorTag Is Nothing Then Exit Sub
    
   
    Dim relative As Variant
    
    relative = RelativeCoordinates(closestFloorTag.insertionPoint, testPoint)
    
    Dim absolute As Variant
    absolute = AddCoordinates(closestFloorTag.insertionPoint, relative)
    
    relative = RelativeCoordinates(closestFloorTag.insertionPoint, testPoint)
    For Each xFloorTag In FloorTags
        absolute = AddCoordinates(xFloorTag.insertionPoint, relative)
        Set Block = ActiveDocument.ModelSpace.InsertBlock(absolute, "FloorConnector.dwg", 1, 1, 1, 0)
        CheckAddLayer "Floor Connectors", PipeColors.Cyan, True, True
        Block.layer = "Floor Connectors"
    Next xFloorTag
    
    'I need one floor tag per elevation box...
    'From there I can figure out which goes where
End Sub
Public Sub SetAttribFromBox(BlockRef As IntelliCAD.blockInsert, AttribName As String, NewValue As String)

    If NewValue = "" Then
        SetAttrib BlockRef, AttribName, 0
    Else
        SetAttrib BlockRef, AttribName, NewValue
    End If
    
End Sub
Public Sub CheckAddBlock()
    Dim Block As AcadBlock
    Set Block = ActiveDocument.Blocks.Add(0, "Test")
    
End Sub
Public Sub ScanPipesAndBreakAtNodes()
Dim Pipe As IntelliCAD.lwPolyline
Dim Block As IntelliCAD.blockInsert
Dim entity As IntelliCAD.entity
Dim nodes As New Collection
Dim pipeIndex As Long
    
    For Each entity In ActiveDocument.ModelSpace
        If entity.EntityName = "BlockInsert" Then
            
            Set Block = entity
            If UCase(Block.Name) = UCase("FakeHead") _
                Or UCase(Block.layer) = UCase("Heads") _
                Or UCase(Block.layer) = UCase("0a__Heads") _
                Or UCase(Block.Name) = UCase("Tee") Then
            
                nodes.Add Block
                
                'CircleLabelBlock block.InsertionPoint, "HH", Colors.vicRed, 0.5, 10
                
            End If
            
        End If
    Next entity
    
    'Exit Sub
    pipeIndex = 0
    For Each entity In ActiveDocument.ModelSpace
        If entity.EntityName = "Polyline" _
            And (UCase(Mid(entity.layer, 1, 6)) = UCase("Pipes.") _
            Or UCase(entity.layer) = UCase("Pipes") _
            Or entity.layer Like "0e__Pipes*") Then
            'Or entity.Layer = "Old Pipes") Then
            
            Set Pipe = entity
            BreakPipe Pipe, nodes
            
            Debug.Print "Pipe " & pipeIndex & " found"
            pipeIndex = pipeIndex + 1
            
        End If
    Next entity
    
End Sub
Public Sub BreakPipe(Pipe As IntelliCAD.lwPolyline, Blocks As Collection)
    'Dim pipe As IntelliCAD.LWPolyline
    Dim Block As IntelliCAD.blockInsert
    Dim entity As IntelliCAD.entity
    
    Dim p As Variant
    Dim ip As IntelliCAD.point
    
    Dim i As Long
    Dim ThisX As Double
    Dim ThisY As Double
    Dim NextX As Double
    Dim NextY As Double
    Dim blockX As Double
    Dim blockY As Double
    Dim blockIsOnLine As Boolean
    Dim pipeStartPoint(0 To 2) As Double
    Dim pipeEndPoint(0 To 2) As Double
    Dim newLine As IntelliCAD.line
    Dim traceLine As IntelliCAD.line
    Dim newCircle As IntelliCAD.Circle
    Dim startCircle As IntelliCAD.Circle
    Dim headsOnSegment As Collection
    Dim vertexesOnPipe As New Collection
    Dim headsOnSegmentLabel As IntelliCAD.text
    Dim segmentStartPoint(0 To 2) As Double
    Dim segmentEndPoint(0 To 2) As Double
    Dim nearestBlock As IntelliCAD.blockInsert
    Dim headsOnSegmentCount As Long
    Dim orderedNonVertextSegmentHeads  As Collection
    Dim headIndex As Long
    Dim hlabel As IntelliCAD.text
    Dim segmentIndex As Long
    Dim newPointIndex As Long
    Dim headOnVertext As Boolean
    Dim headOnNextVertext As Boolean
    Dim newPipe As IntelliCAD.lwPolyline
    Dim newPipeVertices() As Double
    ReDim Preserve newPipeVertices(0 To 1)
    Dim newPipeVerticeBound As Long
    Dim isLastSegment As Boolean
    Dim PipeWidth As Double
    
    'Set Blocks = New Collection
    
    headOnVertext = False
    segmentIndex = 0
    newPointIndex = 0
    isLastSegment = False
    
    pipeStartPoint(0) = Pipe.Coordinates(0)
    pipeStartPoint(1) = Pipe.Coordinates(1)
    pipeStartPoint(2) = Pipe.Coordinates(2)
    pipeEndPoint(0) = Pipe.Coordinates((Pipe.Coordinates.Count - 1) - 2)
    pipeEndPoint(1) = Pipe.Coordinates((Pipe.Coordinates.Count - 1) - 1)
    pipeEndPoint(2) = Pipe.Coordinates((Pipe.Coordinates.Count - 1) - 0)
    newPipeVertices(0) = pipeStartPoint(0)
    newPipeVertices(1) = pipeStartPoint(1)
    newPipeVerticeBound = 1
    
    Pipe.GetWidth 0, PipeWidth, PipeWidth
    
    'CircleLabelBlock pipeStartPoint, "Start", Colors.vicRed, 0.05
    
    For i = 0 To (Pipe.Coordinates.Count - 1) Step 2
            Debug.Print Pipe.Coordinates(i) & ", " & Pipe.Coordinates(i + 1)
        Set headsOnSegment = New Collection
        segmentStartPoint(0) = Pipe.Coordinates(i)
        segmentStartPoint(1) = Pipe.Coordinates(i + 1)
        segmentStartPoint(2) = 0
        headOnVertext = False
        Set orderedNonVertextSegmentHeads = New Collection
    
        ThisX = Pipe.Coordinates(i)
        ThisY = Pipe.Coordinates(i + 1)
        
        'CircleLabelBlock segmentStartPoint, "Segment " & segmentIndex, Colors.vicMagenta, 0.17, 30
        
        'Is there another vertex point after this?
        If (Pipe.Coordinates.Count - 1) >= i + 2 Then
        
            'Setup Next segment point
            NextX = Pipe.Coordinates(i + 2)
            NextY = Pipe.Coordinates(i + 3)
            segmentEndPoint(0) = Pipe.Coordinates(i + 2)
            segmentEndPoint(1) = Pipe.Coordinates(i + 3)
            segmentEndPoint(2) = 0
            
            'DrawTraceLine ThisX, ThisY, NextX, NextY, -0.1, -0.1, vicYellow 'Trace entire polyline with yellow lines
            
            
            
            For Each Block In Blocks
                blockX = Block.insertionPoint.x
                blockY = Block.insertionPoint.y
    
                blockIsOnLine = LinePointXY(ThisX, ThisY, NextX, NextY, blockX, blockY)
                If blockIsOnLine Then
                    'CircleLabelBlock block.InsertionPoint, "HOSeg", Colors.vicRed, 0.5
                    headsOnSegment.Add Block
                End If
            Next Block

        Else
            segmentEndPoint(0) = segmentStartPoint(0)
            segmentEndPoint(1) = segmentStartPoint(1)
            segmentEndPoint(2) = segmentStartPoint(2)
            isLastSegment = True
        End If
        
'        Set headsOnSegmentLabel = ActiveDocument.ModelSpace.AddText("Heads = " & headsOnSegment.Count, segmentStartPoint, 0.1)
'        headsOnSegmentLabel.Layer = "TEST"
        
        headsOnSegmentCount = headsOnSegment.Count
        
        
        'Are there heads on this segment?
        If headsOnSegmentCount > 0 Then 'Yes: there are heads on this segment
            'Process heads in order on the line segment
            For headIndex = 1 To headsOnSegmentCount
                Set nearestBlock = GetRemoveNearestBlock(segmentStartPoint, headsOnSegment)
                If PointsMatch(nearestBlock.insertionPoint, segmentStartPoint) Then
                    'This head is on a vertex, skip it
                    headOnVertext = True
                ElseIf PointsMatch(nearestBlock.insertionPoint, segmentEndPoint) Then
                    'This head is on the next vertex
                    'this is important only on the last segment
                Else
                    'This head is not on a vertex
                    'CircleLabelBlock nearestBlock.InsertionPoint, "H" & headIndex, Colors.vicGreen, 0.2
                    orderedNonVertextSegmentHeads.Add nearestBlock
                    'This is the end of a pipe/polyline, and the beginning of a new pipe/polyline
                End If
                
            Next headIndex
            
            vertexesOnPipe.Add nearestBlock.insertionPoint
            If vertexesOnPipe.Count > 0 Then
                Debug.Print "draw a poly line to the next vertex"
                'CircleLabelBlock segmentStartPoint, "End Vertex " & headIndex, Colors.vicCyan, 0.15

            End If
            Set vertexesOnPipe = New Collection
            
        Else 'No: There are no heads on this segment

            vertexesOnPipe.Add segmentStartPoint
        End If
        
        If headOnVertext Or isLastSegment Then
            'CircleLabelBlock segmentStartPoint, "Point " & newPointIndex & " HoV", Colors.vicMagenta, 0.12, DegreesToRadians(15)
            newPointIndex = newPointIndex + 1
 
            AddNewVertex newPipeVertices, segmentStartPoint
            DrawNewPipe Pipe, newPipeVertices, segmentStartPoint, PipeWidth
            
        Else
            'CircleLabelBlock segmentStartPoint, "Point " & newPointIndex & " NH", Colors.vicMagenta, 0.12, DegreesToRadians(15)
            newPointIndex = newPointIndex + 1

            AddNewVertex newPipeVertices, segmentStartPoint
        End If
        
        If orderedNonVertextSegmentHeads.Count > 0 Then
            For Each Block In orderedNonVertextSegmentHeads
                'CircleLabelBlock block.InsertionPoint, "Point " & newPointIndex & " H", Colors.vicMagenta, 0.12, DegreesToRadians(15)
                newPointIndex = newPointIndex + 1

                AddNewVertex newPipeVertices, Block.insertionPoint
                DrawNewPipe Pipe, newPipeVertices, Block.insertionPoint, PipeWidth
            Next Block
        End If
        
        segmentIndex = segmentIndex + 1
    Next i
    
    Dim oldPipeLayer As IntelliCAD.layer
    Set oldPipeLayer = CheckAddLayer("Old Pipes", Colors.vicWhite)
    Pipe.layer = "Old Pipes"
    oldPipeLayer.LayerOn = False
    
End Sub

Public Function GetRemoveNearestBlock(FromPoint As Variant, Blocks As Collection) As IntelliCAD.blockInsert
Dim closestBlockSoFar As IntelliCAD.blockInsert
Dim closestDistanceSoFar As Double
Dim Block As IntelliCAD.blockInsert
Dim thisBlockDist As Double
Dim closestBlockIndex As Long
Dim index As Long
    
    If Blocks.Count > 0 Then
        thisBlockDist = MeasureBetween(FromPoint, Blocks(1).insertionPoint)
        closestDistanceSoFar = MeasureBetween(FromPoint, Blocks(1).insertionPoint)
        Set closestBlockSoFar = Blocks(1)
        index = 1
        For Each Block In Blocks
            thisBlockDist = MeasureBetween(FromPoint, Block.insertionPoint)
            If thisBlockDist <= closestDistanceSoFar Then
                closestDistanceSoFar = thisBlockDist
                Set closestBlockSoFar = Block
                closestBlockIndex = index
            End If
            index = index + 1
        Next Block
        
        
    End If
    Blocks.Remove closestBlockIndex
    Set GetRemoveNearestBlock = closestBlockSoFar
End Function
Public Sub PrintDoubleArray(InputArray As Variant)
Dim x As Long
Dim str As String
    str = ""
    If Not IsArray(InputArray) Then Exit Sub
    If (InputArray.Count - 1) >= 0 Then
        
        For x = 0 To (InputArray.Count - 1)
            If str = "" Then
                str = InputArray(x)
            Else
                str = str & ", " & InputArray(x)
            End If
        Next x
        
    
    End If
    
    Debug.Print str
End Sub


Public Function GetLineY(Ax As Double, Ay As Double, Bx As Double, By As Double, TestX As Double) As Variant
Dim Slope As Double
Dim Result As Double
Dim Offset As Double

    If (Ax - Bx) = 0 Then
        If (Ay - By) = 0 Then
            'This is a point, not a line, the point has to be the same as A or B
            If TestX = Ax Then
                GetLineY = Ay
            Else
                GetLineY = "Error: Different point!"
            End If
            Exit Function
        Else
            'This is a vertical line, Y is always infinity
            If TestX = Ax Then
                GetLineY = "Infinity"
            Else
                GetLineY = "Error: Point not on line!"
            End If
            Exit Function
        End If
    End If
    
    
    Slope = ((Ay - By) / (Ax - Bx))
    Offset = Ay - Slope * Ax
    Result = Slope * TestX + Offset
    
    
    GetLineY = Result
    

End Function
Public Function InLineBox(Ax As Double, Ay As Double, Bx As Double, By As Double, TestX As Double, TestY As Double) As Double
    If TestX <= Max(Ax, Bx) And TestX >= Min(Ax, Bx) And TestY <= Max(Ay, By) And TestY >= Min(Ay, By) Then
        InLineBox = True
    Else
        InLineBox = False
    End If
End Function
Public Function Max(a As Double, b As Double) As Double
    
    Max = a
    If b > a Then Max = b
    
End Function
Public Function Min(a As Double, b As Double) As Double
    
    Min = a
    If b < a Then Min = b
    
End Function
Public Function LinePointXY(Ax As Double, Ay As Double, Bx As Double, By As Double, TestX As Double, TestY As Double) As Boolean
Dim vYResult As Variant
Dim dblYResult As Double
    
    If InLineBox(Ax, Ay, Bx, By, TestX, TestY) Then
        vYResult = GetLineY(Ax, Ay, Bx, By, TestX)
        
        If vYResult = "Infinity" Then
            dblYResult = TestY  'just make it the same as the test point, any Y value is valid
        ElseIf VarType(vYResult) = vbDouble Then
            dblYResult = vYResult
        Else
            Err.Raise "The GetLineY returned: " & vYResult
        End If
        
        If Abs(dblYResult - TestY) < LINE_POINT_TOLERENCE Then
            LinePointXY = True
        Else
            LinePointXY = False
        End If
    Else
        LinePointXY = False
    End If
    
End Function

Public Function LinePoint(startPoint As Variant, endPoint As Variant, testPoint As Variant) As Boolean
    If (startPoint.Count - 1) <> 1 Then Exit Function
    If (endPoint.Count - 1) <> 1 Then Exit Function
    If (testPoint.Count - 1) < 1 Then Exit Function
    
    LinePoint = LinePointXY(startPoint(0), startPoint(1), endPoint(0), endPoint(1), testPoint(0), testPoint(1))
    
End Function
Public Sub AddNewVertex(VertexArray As Variant, NewPoint As Variant)
Dim UpperBound As Long

    UpperBound = (VertexArray.Count - 1) + 2 'UpperBound + 2
    ReDim Preserve VertexArray(0 To UpperBound)
    VertexArray(UpperBound - 1) = NewPoint(0) 'block.InsertionPoint.x
    VertexArray(UpperBound - 0) = NewPoint(1) 'block.InsertionPoint.y

End Sub
Public Sub RemoveZeroLengthPipes()
Dim Pipe As IntelliCAD.lwPolyline
Dim Block As IntelliCAD.blockInsert
Dim entity As IntelliCAD.entity
Dim Pipes As New Collection
Dim pipeIndex As Long
    
    For Each entity In ActiveDocument.ModelSpace
        If entity.EntityName = "Polyline" Then
            Set Pipe = entity
            If UCase(Mid(Pipe.layer, 1, 6)) = UCase("Pipes.") Or UCase(Pipe.layer) = UCase("Pipes") Then
                
                If Measure(Pipe) = 0 Then
                    Pipe.Delete
'                    Pipes.Add pipe
                    
                    Debug.Print "found one"
                
                End If
            End If
        End If
    Next entity
    

End Sub
Public Sub FindNode()
    Dim xHead As Head
    Dim xTee As Tee
    Dim xRiser As Riser
    
    Dim NodeName As String
    Dim bNodeFound As Boolean
    
    If Heads.Count = 0 Then
        ScanEntities
        MatchEntities
    End If
    
    NodeName = InputBox("Enter node name to find:", "FindNode()", "H.1")
    If NodeName <> "" Then
        For Each xHead In Heads
            If xHead.Name = NodeName Then
                ActiveDocument.Application.ZoomCenter xHead.insertionPoint, 500
                bNodeFound = True
                Exit For
            End If
        Next xHead
        If bNodeFound Then Exit Sub
        For Each xTee In Tees
            If xTee.Name = NodeName Then
                ActiveDocument.Application.ZoomCenter xTee.insertionPoint, 500
                bNodeFound = True
                Exit For
            End If
        Next xTee
        If bNodeFound Then Exit Sub
        For Each xRiser In Risers
            If xRiser.Name = NodeName Then
                ActiveDocument.Application.ZoomCenter xRiser.insertionPoint, 500
                bNodeFound = True
                Exit For
            End If
        Next xRiser
        If Not bNodeFound Then MsgBox "Node " & NodeName & " not found!", vbInformation, "FindNode()"
    End If
    
End Sub
Public Sub FindZeroLengthPipe()
Dim xPipe As Pipe

    For Each xPipe In Pipes
        If xPipe.GetLength() = 0 Then
            '2014'ErrorCircle xPipe.polyline.Coordinates, Colors.vicRed, "Zero Length Pipe", 13, 23, 123
        End If
    Next xPipe
    
End Sub
Public Function DeleteZeroLengthPipes()
Dim xPipe As Pipe
Dim deleteCount As Long
Dim pipeIndex As Long
'Dim deletePipes As New Collection
'Dim deleteMe As Boolean

    deleteCount = 0

    If Pipes Is Nothing Then ScanAndMatchEntities
    If Pipes.Count < 1 Then ScanAndMatchEntities

    For Each xPipe In Pipes
        'deleteMe = False
        If xPipe.GetLength() = 0 Then
            xPipe.polyline.Delete
            'deleteMe = True
            'deletePipes.Add deleteMe
            deleteCount = deleteCount + 1
        End If
    Next xPipe
    
    
    
    DeleteZeroLengthPipes = deleteCount
    
End Function
Public Function DeleteZeroLengthPipesFast()
Dim xPipe As Pipe
Dim deleteCount As Long
Dim pipeIndex As Long
'Dim deletePipes As New Collection
'Dim deleteMe As Boolean
Dim ent As IntelliCAD.entity
Dim polyline As IntelliCAD.lwPolyline


    deleteCount = 0



    For Each ent In ActiveDocument.ModelSpace
        If UCase(Mid(ent.layer, 1, 6)) = UCase("Pipes.") Or UCase(ent.layer) = UCase("Pipes") Then
            If ent.EntityName = "Polyline" Then
                Set polyline = ent
                If Measure(polyline) = 0 Then
                    polyline.Delete
        
                    deleteCount = deleteCount + 1
                End If
            End If
        End If
    Next ent
    
    
    
    DeleteZeroLengthPipesFast = deleteCount
    
End Function
Public Function DeleteOldPipes()
Dim xPipe As Pipe
Dim deleteCount As Long
Dim pipeIndex As Long
'Dim deletePipes As New Collection
'Dim deleteMe As Boolean
Dim ent As IntelliCAD.entity
Dim polyline As IntelliCAD.lwPolyline


    deleteCount = 0



    For Each ent In ActiveDocument.ModelSpace
        If ent.layer = "Old Pipes" Then
            If ent.EntityName = "Polyline" Then
                Set polyline = ent
                'If Measure(polyline) = 0 Then
                    polyline.Delete
        
                    deleteCount = deleteCount + 1
                'End If
            End If
        End If
    Next ent
    
    
    
    DeleteOldPipes = deleteCount
    
End Function
Public Sub RunDeleteZeroLengthPipes(Optional showMessages As Variant)
Dim deleteCount As Long
Dim deleteOldPipesCount As Long
Dim cloneCount As Long
    
    If IsMissing(showMessages) Then showMessages = True
    
    deleteCount = DeleteZeroLengthPipesFast()
    
    cloneCount = RemovePipeClones
    
    If showMessages Then
        MsgBox deleteCount & " zero length pipes were deleted!" & vbCrLf & _
            cloneCount & " clones were deleted!", vbInformation, "RunDeleteZeroLengthPipes()"
    Else
        Debug.Print "RunDeleteZeroLengthPipes: " & deleteCount & " zero length pipes were deleted!" & vbCrLf & _
            cloneCount & " clones were deleted!", vbInformation, "RunDeleteZeroLengthPipes()"
    End If
    deleteOldPipesCount = DeleteOldPipes()
    If showMessages Then
        MsgBox deleteCount & " old pipes were deleted!", vbInformation, "RunDeleteZeroLengthPipes()"
    Else
        Debug.Print "RunDeleteZeroLengthPipes: " & deleteCount & " old pipes were deleted!", vbInformation, "RunDeleteZeroLengthPipes()"
    End If
End Sub

Public Function GetHeadModelSlopes(HeadModelNumber As String) As Collection
Dim xHeadModel As HeadModel
Dim xSlope As HeadModelSlope
Dim xHeadModelItem As HeadModelItem

    If headModels Is Nothing Then
        basData.LoadHeadModels
    End If
    
    Set GetHeadModelSlopes = headModels(HeadModelNumber).Slopes
End Function
Public Function CoverageIsSidewall(Coverage As String) As Boolean
Dim a As String
Dim b As String
    If Len(Coverage) = 5 Then
        a = Mid(Coverage, 1, 2)
        b = Mid(Coverage, 4, 2)
        If a <> b Then
            CoverageIsSidewall = True
        Else
            CoverageIsSidewall = False
        End If
    Else
        CoverageIsSidewall = False
        
    End If
End Function
Public Function ModelCodeIsSidewall(ModelCode As String) As Boolean
Dim a As String
Dim b As String
Dim swIndex As Long
    If Len(ModelCode) >= 2 Then
        swIndex = InStr(1, ModelCode, "SW")
        If swIndex > 0 Then
            ModelCodeIsSidewall = True
        Else
            ModelCodeIsSidewall = False
        End If
    Else
        ModelCodeIsSidewall = False
    End If
End Function
Public Sub InsertHeadModelCode(Coverage As String, side As String, ModelCode As String)
On Error GoTo ErrTrap
    Dim startPoint As Variant
    Dim blockHead As IntelliCAD.blockInsert
    Dim strBlockFileName As String
    Dim strSendCommand As String
    strBlockFileName = "Head.dwg"
    
    If Len(side) > 0 Then side = Mid(side, 1, 1)
    
    If ModelCodeIsSidewall(ModelCode) Then
        strSendCommand = "(insertsidewallheadmodelcode """ & Coverage & """ """ & side & """ """ & ModelCode & """ ) "
    Else
        strSendCommand = "(insertheadmodelcode """ & Coverage & """ """ & ModelCode & """ ) "
    End If
    
    
DoAgain:
    
    ActiveActiveDocument.SendCommand strSendCommand
    ActiveActiveDocument.Activate
    
'    GoTo DoAgain

ExitProc:
    Exit Sub

ErrTrap:
    If Err.Number = -2147352567 Then 'The user pressed escape
        GoTo ExitProc
    ElseIf Err.Number = -2145320928 Then  'user pressed enter
        GoTo ExitProc
    ElseIf Err.Number = -2145386445 Then 'Can't find block file
        MsgBox "Cannot find the block file " & strBlockFileName, vbCritical, "InsertHeadModel()"
        GoTo ExitProc
    Else
        MsgBox Err.Number & " " & Err.Description, vbCritical, "InsertHeadModel()"
        GoTo ExitProc
    End If
End Sub

Public Function GetCoverageDepth(Coverage As String)
Dim x As Long

    x = InStr(1, Coverage, "x")
    If x > 1 And x < Len(Coverage) Then
        GetCoverageDepth = Mid(Coverage, 1, x - 1)
    Else
        Err.Raise 1, "GetCoverageDepth(Coverage)", "There coverage string is not in the format 12x14."
    End If
    
End Function

Public Function GetCoverageWidth(Coverage As String)
Dim x As Long

    x = InStr(1, Coverage, "x")
    If x > 1 And x < Len(Coverage) Then
        GetCoverageWidth = Mid(Coverage, x + 1)
    Else
        Err.Raise 1, "GetCoverageWidth(Coverage)", "There is 'x' character in the coverage, should be in the format 12x14."
    End If
    
End Function
Public Function ScanLineEndsAlmost()
Dim xPipe As Pipe
Dim xNode As Node
Dim xHeadPair As HeadPair
Dim lngStartAlmosts As Long
Dim lngEndAlmosts As Long
Dim dblDistance As Double
Const MAX_DIST = 0.2
    
    'Match Nodes with pipe ends
    ScanAndMatchEntities
    DeleteErrorCircles
    
    lngStartAlmosts = 0
    lngEndAlmosts = 0
    
    For Each xPipe In Pipes
        For Each xNode In nodes
            dblDistance = MeasureBetween(xNode.insertionPoint, xPipe.startPoint)
            If dblDistance < MAX_DIST And dblDistance > 0 Then
                xPipe.SetStartPoint xNode.insertionPoint
                lngStartAlmosts = lngStartAlmosts + 1
            End If
            dblDistance = MeasureBetween(xNode.insertionPoint, xPipe.endPoint)
            If dblDistance < MAX_DIST And dblDistance > 0 Then
                xPipe.SetEndPoint xNode.insertionPoint
                lngEndAlmosts = lngEndAlmosts + 1
            End If
        Next xNode
    Next xPipe
    
    For Each xHeadPair In HeadPairs
        For Each xNode In nodes
            dblDistance = MeasureBetween(xNode.insertionPoint, xHeadPair.HeadPairLine.startPoint)
            If dblDistance < MAX_DIST And dblDistance > 0 Then
                'xPipe.SetStartPoint xNode.InsertionPoint
                xHeadPair.HeadPairLine.startPoint = xNode.insertionPoint
                lngStartAlmosts = lngStartAlmosts + 1
            End If
            dblDistance = MeasureBetween(xNode.insertionPoint, xHeadPair.HeadPairLine.endPoint)
            If dblDistance < MAX_DIST And dblDistance > 0 Then
                'xPipe.SetEndPoint xNode.InsertionPoint
                xHeadPair.HeadPairLine.endPoint = xNode.insertionPoint
                lngEndAlmosts = lngEndAlmosts + 1
            End If
        Next xNode
    Next xHeadPair
    
    Debug.Print "StartAlmosts = " & lngStartAlmosts
    Debug.Print "EndAlmosts = " & lngEndAlmosts

End Function
Public Function TestScanLineEnds()
Dim xPipe As Pipe
Dim xNode As Node
Dim lngStartAlmosts As Long
Dim lngEndAlmosts As Long
Dim dblStartDistance As Double
Dim dblEndDistance As Double
Dim polyCount As Long
Dim newCoords() As Double
Dim x As Long
    
    'Match Nodes with pipe ends
    
    lngStartAlmosts = 0
    lngEndAlmosts = 0
    Dim entity As IntelliCAD.entity
    Dim points As Variant
    Dim poly As IntelliCAD.lwPolyline
    ActiveDocument.Utility.GetEntity entity, points, "Select Polyline"
    
    Debug.Print "Got one"
    If entity.EntityName = "Polyline" Then
        Debug.Print "Got a polyline"
        Set poly = entity
        polyCount = (poly.Coordinates.Count - 1)
        ReDim newCoords(polyUBound)
        For x = 0 To polyUBound
        
            newCoords(x) = poly.Coordinates(x)
        Next x
        
        
        newCoords(0) = 0
        newCoords(1) = 0
        poly.Coordinates = newCoords
        poly.Update
    End If
    
    
    
End Function
Public Sub MeasurePipeExceptRisers()
Dim xPipe As Pipe
Dim PipeTypes As New Collection
Dim pType As PipeType
Dim vPipeSize As Variant

    For Each xPipe In Pipes
        If CollectionContainsNameString(PipeTypes, CStr(xPipe.Diameter)) Then
           Set pType = PipeTypes(CStr(xPipe.Diameter))
           pType.TotalLength = pType.TotalLength + xPipe.ActualLength
        Else
           Set pType = New PipeType
           pType.InternalDiameter = xPipe.Diameter
           pType.Name = CStr(xPipe.Diameter)
           pType.TotalLength = xPipe.ActualLength
           PipeTypes.Add pType, CStr(pType.Name)
        End If
    Next xPipe
    
    For Each pType In PipeTypes
        Debug.Print pType.Name & " = " & pType.TotalLength / 12
    Next pType
End Sub
Public Function RemovePipeClones() As Long
Dim xPipe As Pipe
Dim innerPipe As Pipe
Dim pipeClone As Pipe
Dim pipeClones As Collection
Dim removePipesIndexes As New Collection
Dim removePipes As New Collection
Dim isFirstClone As Boolean
Dim xPipeIndex As Variant
Dim clonesToRemoveCount As Long

    If Pipes Is Nothing Then
        MsgBox "There are no pipes!", vbCritical, "Scanner.RemovePipeClones()"
        Exit Function
    End If
    
    xPipeIndex = 1
    For Each xPipe In Pipes
        If Not xPipe.polyline Is Nothing Then
            If (xPipe.polyline.Coordinates.Count - 1) > 0 Then
                Set pipeClones = GetPipeClones(xPipe, Pipes)
                Debug.Print "RemovePipeClones: Pipe" & xPipe.PipeNumber & " Clones " & pipeClones.Count & " Count-1 = " & (xPipe.polyline.Coordinates.Count - 1)
                If pipeClones.Count > 1 Then
                    
                    isFirstClone = True
                    For Each pipeClone In pipeClones
                        If Not isFirstClone And pipeClone.PipeNumber = xPipe.PipeNumber Then    'Adds only clones that aren't the first one.
                            removePipesIndexes.Add xPipeIndex
                            removePipes.Add xPipe
                        End If
                        isFirstClone = False
                    Next pipeClone
                    
                End If
            End If
        End If
        xPipeIndex = xPipeIndex + 1
    Next xPipe
    clonesToRemoveCount = removePipes.Count
    For Each xPipe In removePipes
        xPipe.polyline.Delete
        Debug.Print "Deleting: Pipe" & xPipe.PipeNumber & " ..."
    Next xPipe
    For xPipeIndex = removePipesIndexes.Count To 1 Step -1
        Pipes.Remove xPipeIndex
    Next xPipeIndex
    RemovePipeClones = clonesToRemoveCount
End Function
Public Function GetPipeClones(checkPipe As Pipe, Pipes As Collection) As Collection
Dim xPipe As Pipe
Dim matchCount As Long
Dim pipeClones As New Collection
    
    For Each xPipe In Pipes
        If MatchCoordinates(xPipe.polyline.Coordinates, checkPipe.polyline.Coordinates) Then
        'If xPipe.Polyline.Coordinates = checkPipe.Polyline.Coordinates Then
            pipeClones.Add xPipe
        End If
    Next xPipe
    
    Set GetPipeClones = pipeClones
    
End Function
Public Function MatchCoordinates(a As Variant, b As Variant) As Boolean
Dim i As Long
    If (a.Count) <> (b.Count) Then
        MatchCoordinates = False
        Exit Function
    End If
    For i = 0 To (a.Count)
        If a(i) <> b(i) Then
            MatchCoordinates = False
            Exit Function
        End If
    Next i
    MatchCoordinates = True
End Function
