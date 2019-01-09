Attribute VB_Name = "DrawingBuilder"
Public Sub Test()
    Dim inputFilePath As String

    inputFilePath = "E:\X-Fire\Test CadDrawingScan.cds" 'VBA.Environ("Temp") & "\CURRENT_DRAWING.xml"
    
    Build inputFilePath
    
End Sub
Public Sub BuildFromTemp()
    
    Build VBA.Environ("Temp") & "\CURRENT_DRAWING.xml"
    
End Sub
Public Sub Build(Optional inputFilePath As Variant)
    Dim objDoc As MSXML2.DOMDocument60
    Dim objNode As MSXML2.IXMLDOMNode
    Dim objRoot As MSXML2.IXMLDOMElement
    Dim documentElem As MSXML2.IXMLDOMElement
    Dim blockInsertsElem As MSXML2.IXMLDOMElement
    Dim lwPolylinesElem As MSXML2.IXMLDOMElement
    Dim polyLinesElem As MSXML2.IXMLDOMElement
    Dim linesElem As MSXML2.IXMLDOMElement
    Dim mTextsElem As MSXML2.IXMLDOMElement
    Dim textsElem As MSXML2.IXMLDOMElement
    Dim imagesElem As MSXML2.IXMLDOMElement
    Dim biNode As MSXML2.IXMLDOMNode
    Dim nameNode As MSXML2.IXMLDOMNode
    Dim biNodes As MSXML2.IXMLDOMNodeList
    
    Dim entity As IntelliCAD.entity
    Dim newDocument As IntelliCAD.Document
    
    Dim loaded As Boolean
    
    Set objDoc = New MSXML2.DOMDocument60
    loaded = objDoc.Load(inputFilePath)
    
    Set newDocument = IntelliCAD.Application.Documents.Add
    newDocument.Activate
    
    If loaded Then
        Debug.Print "Loaded"
        
        objDoc.setProperty "SelectionLanguage", "XPath"
        BuildTextStyles objDoc
        BuildLayers objDoc
        BuildBlocks objDoc
        
        ActiveDocument.Application.ZoomExtents
    Else
        Debug.Print "Not loaded"
    End If

End Sub
Public Sub BuildLWPolylines(node As MSXML2.IXMLDOMNode, block As IntelliCAD.block)
    Dim objDoc As MSXML2.DOMDocument60
    Dim lineNodes As MSXML2.IXMLDOMNodeList
    Dim lineNode As MSXML2.IXMLDOMNode
    Dim pointNodes As MSXML2.IXMLDOMNodeList
    Dim segmentNodes As MSXML2.IXMLDOMNodeList
    Dim pointNode As MSXML2.IXMLDOMNode
    Dim segmentNode As MSXML2.IXMLDOMNode
    
    Dim points As IntelliCAD.points
    Dim lwPolyline As IntelliCAD.lwPolyline
        
    Set objDoc = node.ownerDocument
        
    Set lineNodes = node.selectNodes("LWPolylines/LWPolyline")
    
    For Each lineNode In lineNodes
        Set points = New IntelliCAD.points
        Set pointNodes = lineNode.selectNodes("Coordinates/Point")
        
        For Each pointNode In pointNodes
            points.Add _
                pointNode.selectSingleNode("@X").text, _
                pointNode.selectSingleNode("@Y").text, _
                pointNode.selectSingleNode("@Z").text
        Next pointNode
        
        Set lwPolyline = block.AddLightWeightPolyline(points)
        lwPolyline.Color.ColorIndex = lineNode.selectSingleNode("@Color").text
        lwPolyline.layer = lineNode.selectSingleNode("@Layer").text
        lwPolyline.Thickness = lineNode.selectSingleNode("@Thickness").text
        lwPolyline.Closed = lineNode.selectSingleNode("@Closed").text
        lwPolyline.LineWeight = lineNode.selectSingleNode("@LineWeight").text
        lwPolyline.LineType = lineNode.selectSingleNode("@LineType").text
         
        Set segmentNodes = lineNode.selectNodes("SegmentWidths/Segment")
        Dim i As Long
        i = 0
        For Each segmentNode In segmentNodes
            lwPolyline.SetWidth _
                i, _
                segmentNode.selectSingleNode("@StartWidth").text, _
                segmentNode.selectSingleNode("@EndWidth").text
                
            i = i + 1
        Next segmentNode
    Next lineNode
    
End Sub
Public Sub BuildAttributeDef(node As MSXML2.IXMLDOMNode, block As IntelliCAD.block)
    Dim objDoc As MSXML2.DOMDocument60
    Dim nodes As MSXML2.IXMLDOMNodeList
    Dim nodeItem As MSXML2.IXMLDOMNode
    Dim attrDefNode As MSXML2.IXMLDOMNode
    Dim insertPointNode As MSXML2.IXMLDOMNode
        
    Dim attrDef As IntelliCAD.AttributeDef
    Dim insertPoint As IntelliCAD.point
    Dim mText As IntelliCAD.mText
    
    Set objDoc = node.ownerDocument
    Set nodes = node.selectNodes("AttributeDefs/AttributeDef")
    
    For Each nodeItem In nodes
        
        Set insertPointNode = nodeItem.selectSingleNode("InsertionPoint")
        Set insertPoint = BuildPoint(insertPointNode)
        
        Set attrDef = block.AddAttributeDef( _
            nodeItem.selectSingleNode("@Height").text, _
            nodeItem.selectSingleNode("@Mode").text, _
            "No Prompt", _
            insertPoint, _
            nodeItem.selectSingleNode("@TagString").text, _
            nodeItem.selectSingleNode("@TextString").text)
        
        attrDef.styleName = nodeItem.selectSingleNode("@StyleName").text
        attrDef.Color.ColorIndex = nodeItem.selectSingleNode("@Color").text
        attrDef.layer = nodeItem.selectSingleNode("@Layer").text
    Next nodeItem
    
End Sub
Public Sub BuildCircle(node As MSXML2.IXMLDOMNode, block As IntelliCAD.block)
    Dim objDoc As MSXML2.DOMDocument60
    Dim circ As IntelliCAD.Circle
    Dim center As IntelliCAD.point
    Dim mText As IntelliCAD.mText
    
    Dim nodes As MSXML2.IXMLDOMNodeList
    Dim nodeItem As MSXML2.IXMLDOMNode
    Dim circleNode As MSXML2.IXMLDOMNode
    Dim centerNode As MSXML2.IXMLDOMNode
        
    Set objDoc = node.ownerDocument
    Set nodes = node.selectNodes("Circles/Circle")
    
    For Each nodeItem In nodes
        
        Set centerNode = nodeItem.selectSingleNode("Center")
        Set center = BuildPoint(centerNode)
        
        Set circ = block.AddCircle( _
            center, _
            nodeItem.selectSingleNode("@Radius").text)
        
        circ.Color.ColorIndex = nodeItem.selectSingleNode("@Color").text
        circ.layer = nodeItem.selectSingleNode("@Layer").text
         
    Next nodeItem
    
End Sub
Public Sub BuildMText(node As MSXML2.IXMLDOMNode, block As IntelliCAD.block)
    Dim objDoc As MSXML2.DOMDocument60
    Dim nodes As MSXML2.IXMLDOMNodeList
    Dim nodeItem As MSXML2.IXMLDOMNode
    Dim pointNode As MSXML2.IXMLDOMNode
    
    Dim point As IntelliCAD.point
    Dim mText As IntelliCAD.mText
    
    Set objDoc = node.ownerDocument
    Set nodes = node.selectNodes("MTexts/MText")
    
    For Each nodeItem In nodes
        
        Set pointNode = nodeItem.selectSingleNode("InsertionPoint")
        Set point = BuildPoint(pointNode)
        
        Set mText = block.AddMText( _
            point, _
            nodeItem.selectSingleNode("@Width").text, _
            nodeItem.selectSingleNode("@TextString").text)
        
        mText.Color.ColorIndex = nodeItem.selectSingleNode("@Color").text
        mText.layer = nodeItem.selectSingleNode("@Layer").text
        mText.Height = nodeItem.selectSingleNode("@Height").text
        mText.Rotation = nodeItem.selectSingleNode("@Rotation").text
         
    Next nodeItem
    
End Sub
Public Sub BuildText(node As MSXML2.IXMLDOMNode, block As IntelliCAD.block)
    Dim objDoc As MSXML2.DOMDocument60
    Dim nodes As MSXML2.IXMLDOMNodeList
    Dim nodeItem As MSXML2.IXMLDOMNode
    Dim pointNode As MSXML2.IXMLDOMNode
        
    Dim point As IntelliCAD.point
    Dim text As IntelliCAD.text
    
    Set objDoc = node.ownerDocument
    Set nodes = node.selectNodes("Texts/Text")
    
    For Each nodeItem In nodes
        Set pointNode = nodeItem.selectSingleNode("InsertionPoint")
        Set point = BuildPoint(pointNode)
        
        Set text = block.AddText( _
            nodeItem.selectSingleNode("@TextString").text, _
            point, _
            CDbl(nodeItem.selectSingleNode("@Height").text))
        
        text.Color.ColorIndex = nodeItem.selectSingleNode("@Color").text
        text.layer = nodeItem.selectSingleNode("@Layer").text
        text.styleName = nodeItem.selectSingleNode("@StyleName").text
        text.HorizontalAlignment = nodeItem.selectSingleNode("@HorizontalAlignment").text
        text.VerticalAlignment = nodeItem.selectSingleNode("@VerticalAlignment").text
        text.ScaleFactor = nodeItem.selectSingleNode("@ScaleFactor").text
    Next nodeItem
    
End Sub
Public Sub BuildBlockInsert(node As MSXML2.IXMLDOMNode, block As IntelliCAD.block)
    Dim objDoc As MSXML2.DOMDocument60
    Dim point As IntelliCAD.point
    Dim blockInsert As IntelliCAD.blockInsert
    
    Dim nodes As MSXML2.IXMLDOMNodeList
    Dim nodeItem As MSXML2.IXMLDOMNode
    Dim pointNode As MSXML2.IXMLDOMNode
    Dim blockInsertElem As MSXML2.IXMLDOMElement
    Dim insertPointElem As MSXML2.IXMLDOMElement
    Dim attribInsertPointElem As MSXML2.IXMLDOMElement
    Dim textInsertPointElem As MSXML2.IXMLDOMElement
    Dim attributesElem As MSXML2.IXMLDOMElement
    Dim attribElem As MSXML2.IXMLDOMElement
    
    Dim blkInsert As IntelliCAD.blockInsert
    Dim attribs As IntelliCAD.Attributes
    Dim attr As IntelliCAD.Attribute
    Dim insertPoint As IntelliCAD.point
    Dim attribInsertPoint As IntelliCAD.point
    Dim textInsertPoint As IntelliCAD.point
    
    Dim attrName As String
    Dim blockName As String
            
    Set objDoc = node.ownerDocument
    Set nodes = node.selectNodes("BlockInserts/BlockInsert")
    
    For Each nodeItem In nodes
        Set pointNode = nodeItem.selectSingleNode("InsertionPoint")
        Set point = BuildPoint(pointNode)
        blockName = nodeItem.selectSingleNode("@Name").text
        
        If block.name = "*Model_Space" Then
            Set blockInsert = ActiveDocument.ModelSpace.InsertBlock( _
            point, _
            blockName, _
            nodeItem.selectSingleNode("@XScaleFactor").text, _
            nodeItem.selectSingleNode("@YScaleFactor").text, _
            nodeItem.selectSingleNode("@ZScaleFactor").text, _
            nodeItem.selectSingleNode("@Rotation").text)
        Else
            Set blockInsert = block.InsertBlock( _
                point, _
                blockName, _
                nodeItem.selectSingleNode("@XScaleFactor").text, _
                nodeItem.selectSingleNode("@YScaleFactor").text, _
                nodeItem.selectSingleNode("@ZScaleFactor").text, _
                nodeItem.selectSingleNode("@Rotation").text)
        End If
        
        blockInsert.Color.ColorIndex = nodeItem.selectSingleNode("@Color").text
        blockInsert.layer = nodeItem.selectSingleNode("@Layer").text
        blockInsert.Rotation = nodeItem.selectSingleNode("@Rotation").text
        blockInsert.Visible = nodeItem.selectSingleNode("@Visible").text
          
          If blockInsert.HasAttributes Then
            Set attribs = blockInsert.GetAttributes()
            For i = 0 To attribs.Count - 1
                Set attr = attribs.Item(i)
                
                attrName = "Attributes/Attribute[@TagString='" & attr.TagString & "']"
                Set attribElem = nodeItem.selectSingleNode(attrName)
                
                If Not attribElem Is Nothing Then
                    attr.TextString = attribElem.selectSingleNode("@TextString").text
                    attr.Height = attribElem.selectSingleNode("@Height").text
                    attr.Rotation = attribElem.selectSingleNode("@Rotation").text
                    attr.ScaleFactor = attribElem.selectSingleNode("@ScaleFactor").text
                    attr.Invisible = attribElem.selectSingleNode("@Invisible").text
                                
                    Set pointNode = attribElem.selectSingleNode("TextAlignmentPoint")
                    attr.TextAlignmentPoint = BuildPoint(pointNode)
                    
                    Set pointNode = attribElem.selectSingleNode("InsertionPoint")
                    attr.insertionPoint = BuildPoint(pointNode)
                End If
            Next i
        End If
    Next nodeItem

End Sub
Public Function BuildPoints(pointNodes As MSXML2.IXMLDOMNodeList) As Collection
    Dim pointNode As MSXML2.IXMLDOMNode
    
    Dim points As New Collection

    For Each pointNode In pointNodes
        points.Add _
            pointNode.selectSingleNode("@X").text, _
            pointNode.selectSingleNode("@Y").text, _
            pointNode.selectSingleNode("@Z").text
    Next pointNode

    Set BuildPoints = points
    
End Function
Public Function BuildPoint(pointNode As MSXML2.IXMLDOMNode) As IntelliCAD.point
    Dim point As New IntelliCAD.point

    point.x = pointNode.selectSingleNode("@X").text
    point.y = pointNode.selectSingleNode("@Y").text
    point.Z = pointNode.selectSingleNode("@Z").text

    Set BuildPoint = point
    
End Function

Public Sub BuildLines(node As MSXML2.IXMLDOMNode, block As IntelliCAD.block)
    Dim objDoc As MSXML2.DOMDocument60
    Dim lineNodes As MSXML2.IXMLDOMNodeList
    Dim lineNode As MSXML2.IXMLDOMNode
    Dim startPointNode As MSXML2.IXMLDOMNode
    Dim endPointNode As MSXML2.IXMLDOMNode
    
    Dim line As IntelliCAD.line
    Dim startPoint As IntelliCAD.point
    Dim endPoint As IntelliCAD.point
    
    Set objDoc = node.ownerDocument
    Set lineNodes = node.selectNodes("Lines/Line")
    
    For Each lineNode In lineNodes
        Set startPointNode = lineNode.selectSingleNode("StartPoint")
        Set startPoint = BuildPoint(startPointNode)
        
        Set endPointNode = lineNode.selectSingleNode("EndPoint")
        Set endPoint = BuildPoint(endPointNode)
        
        Set line = block.AddLine(startPoint, endPoint)
        line.Color.ColorIndex = lineNode.selectSingleNode("@Color").text
        line.layer = lineNode.selectSingleNode("@Layer").text
        line.Thickness = lineNode.selectSingleNode("@Thickness").text
        line.LineWeight = lineNode.selectSingleNode("@LineWeight").text
        line.LineType = lineNode.selectSingleNode("@LineType").text
        line.Visible = lineNode.selectSingleNode("@Visible").text
    Next lineNode
    
End Sub

Public Sub BuildTextStyles(objDoc As MSXML2.DOMDocument60)
    Dim styleNodes As MSXML2.IXMLDOMNodeList
    Dim styleNode As MSXML2.IXMLDOMNode
    
    Dim textStyle As IntelliCAD.textStyle
    Dim styleName As String
        
    Set styleNodes = objDoc.selectNodes("/CadData/TextStyles/TextStyle")
    
    For Each styleNode In styleNodes
        styleName = styleNode.selectSingleNode("@Name").text
        If styleName <> "Standard" And Trim(styleName) <> "" Then
            Set textStyle = ActiveDocument.TextStyles.Add(styleName)
            'textStyle.Color.ColorIndex = styleNode.selectSingleNode("@Color").text
            textStyle.Height = styleNode.selectSingleNode("@Height").text
            textStyle.Width = styleNode.selectSingleNode("@Width").text
            textStyle.FontFile = styleNode.selectSingleNode("@FontFile").text
            textStyle.BigFontFile = styleNode.selectSingleNode("@BigFontFile").text
        End If
    Next styleNode
    
End Sub

Public Sub BuildLayers(objDoc As MSXML2.DOMDocument60)
    Dim layerNodes As MSXML2.IXMLDOMNodeList
    Dim layerNode As MSXML2.IXMLDOMNode
    
    Dim layer As IntelliCAD.layer
        
    Set layerNodes = objDoc.selectNodes("/CadData/Layers/Layer")
    
    For Each layerNode In layerNodes
        Set layer = ActiveDocument.Layers.Add(layerNode.selectSingleNode("@Name").text)
        
        layer.Color.SetRGB _
            layerNode.selectSingleNode("@ColorRed").text, _
            layerNode.selectSingleNode("@ColorGreen").text, _
            layerNode.selectSingleNode("@ColorBlue").text
            
        layer.Color.EntityColor = layerNode.selectSingleNode("@EntityColor").text
        layer.Color.ColorMethod = layerNode.selectSingleNode("@ColorMethod").text
        layer.Color.ColorIndex = layerNode.selectSingleNode("@ColorRed").text
        layer.LineWeight = layerNode.selectSingleNode("@LineWeight").text
        layer.Freeze = layerNode.selectSingleNode("@Freeze").text
        layer.Lock = layerNode.selectSingleNode("@Lock").text
        layer.LayerOn = layerNode.selectSingleNode("@LayerOn").text
        layer.Plottable = layerNode.selectSingleNode("@Plottable").text
        layer.PlotStyleName = layerNode.selectSingleNode("@PlotStyleName").text
    Next layerNode
    
End Sub
Public Sub BuildBlocks(objDoc As MSXML2.DOMDocument60)
    Dim blockNodes As MSXML2.IXMLDOMNodeList
    Dim blockNode As MSXML2.IXMLDOMNode
    Dim modelSpaceBlockNode As MSXML2.IXMLDOMNode
    Dim originNode As MSXML2.IXMLDOMNode
    
    Dim block As IntelliCAD.block
    Dim modelSpaceBlock As IntelliCAD.block
    Dim insertionPoint As IntelliCAD.point
    Dim blockName As String
        
    Set blockNodes = objDoc.selectNodes("/CadData/Blocks/Block")
    
    For Each blockNode In blockNodes
        Set originNode = blockNode.selectSingleNode("Origin")
        Set insertionPoint = BuildPoint(originNode)
        blockName = blockNode.selectSingleNode("@Name").text
        If blockName = "*Model_Space" Then
            Set modelSpaceBlock = ActiveDocument.Blocks.Item(blockName)
            Set modelSpaceBlockNode = blockNode
        ElseIf Not blockName Like "[*]*" Then
            Set block = ActiveDocument.Blocks.Add(insertionPoint, blockName)
        Else
            Set block = Nothing
        End If
        
        If Not block Is Nothing Then
            BuildLWPolylines blockNode, block
            BuildMText blockNode, block
            BuildText blockNode, block
            BuildLines blockNode, block
            BuildCircle blockNode, block
            BuildAttributeDef blockNode, block
            BuildBlockInsert blockNode, block
        End If
        
    Next blockNode
    
    If Not modelSpaceBlock Is Nothing Then
        BuildLWPolylines modelSpaceBlockNode, modelSpaceBlock
        BuildMText modelSpaceBlockNode, modelSpaceBlock
        BuildText modelSpaceBlockNode, modelSpaceBlock
        BuildLines modelSpaceBlockNode, modelSpaceBlock
        BuildCircle modelSpaceBlockNode, modelSpaceBlock
        BuildAttributeDef modelSpaceBlockNode, modelSpaceBlock
        BuildBlockInsert modelSpaceBlockNode, modelSpaceBlock
    End If
    
End Sub




