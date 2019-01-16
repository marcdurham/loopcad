Attribute VB_Name = "DrawingScanner"
Option Explicit
Public Sub Test()
    Scan
    '"F:\Test.cadx"
End Sub
Public Sub ScanToTemp()
    Scan VBA.Environ("Temp") & "\CURRENT_DRAWING.xml"
End Sub


Public Sub Scan(Optional outputFilePath As Variant)
    Dim ObjDoc As MSXML2.DOMDocument60
    Dim objNode As MSXML2.IXMLDOMNode
    Dim ObjRoot As MSXML2.IXMLDOMElement
    Dim documentElem As MSXML2.IXMLDOMElement
    Dim layersElem As MSXML2.IXMLDOMElement
    Dim blocksElem As MSXML2.IXMLDOMElement
    Dim blockElem As MSXML2.IXMLDOMElement
    Dim blockInsertsElem As MSXML2.IXMLDOMElement
    Dim lwPolylinesElem As MSXML2.IXMLDOMElement
    Dim polyLinesElem As MSXML2.IXMLDOMElement
    Dim linesElem As MSXML2.IXMLDOMElement
    Dim mTextsElem As MSXML2.IXMLDOMElement
    Dim textsElem As MSXML2.IXMLDOMElement
    Dim imagesElem As MSXML2.IXMLDOMElement
    
    Dim layer As IntelliCAD.layer
    Dim Block As IntelliCAD.Block
    Dim entity As IntelliCAD.entity
    
    If IsMissing(outputFilePath) Then
        outputFilePath = ShowSave("Save Scan File As...", "CadDrawingScan", "Cad Drawing Scan (*.cds)", "*.cds", ActiveDocument.path)
        If outputFilePath = "" Then
            Exit Sub
        End If
    End If
    
    Set ObjDoc = New DOMDocument60
    ObjDoc.resolveExternals = True
    
    Set objNode = ObjDoc.createProcessingInstruction("xml", "version='1.0' encoding='UTF-8'")
    Set objNode = ObjDoc.insertBefore(objNode, ObjDoc.childNodes.Item(0))
    Set ObjRoot = ObjDoc.createElement("CadData")
    Set ObjDoc.documentElement = ObjRoot
    
    Set documentElem = ObjDoc.createElement("Document")
    ObjRoot.appendChild documentElem
    documentElem.setAttribute "FullName", ActiveDocument.FullName
    documentElem.setAttribute "Name", ActiveDocument.Name
    documentElem.setAttribute "Path", ActiveDocument.path
    
    Set layersElem = ObjDoc.createElement("Layers")
    ObjRoot.appendChild layersElem
    
    For Each layer In ActiveDocument.Layers
        CheckLayers layer, ObjDoc, layersElem
    Next layer
    
    Dim textStylesElem As MSXML2.IXMLDOMElement
    Dim txtSty As IntelliCAD.textStyle
    Set textStylesElem = ObjDoc.createElement("TextStyles")
    ObjRoot.appendChild textStylesElem
    
    For Each txtSty In ActiveDocument.TextStyles
        CheckTextStyles txtSty, ObjDoc, textStylesElem
    Next txtSty
    
    Set blocksElem = ObjDoc.createElement("Blocks")
    ObjRoot.appendChild blocksElem
    
    For Each Block In ActiveDocument.Blocks
        Set blockElem = ObjDoc.createElement("Block")
        blocksElem.appendChild blockElem
        blockElem.setAttribute "Name", Block.Name
        blockElem.setAttribute "IsLayout", Block.IsLayout
        blockElem.setAttribute "IsXRef", Block.IsXRef
        
        ScanBlock Block, ObjDoc, blockElem
    Next Block
    
    ObjDoc.Save outputFilePath

End Sub
Public Sub ScanBlock(Block As IntelliCAD.blockInsert, ObjDoc As MSXML2.DOMDocument60, ObjRoot As MSXML2.IXMLDOMElement)
    Dim objNode As MSXML2.IXMLDOMNode
    Dim documentElem As MSXML2.IXMLDOMElement
    Dim layersElem As MSXML2.IXMLDOMElement
    Dim blocksElem As MSXML2.IXMLDOMElement
    Dim blockInsertsElem As MSXML2.IXMLDOMElement
    Dim lwPolylinesElem As MSXML2.IXMLDOMElement
    Dim polyLinesElem As MSXML2.IXMLDOMElement
    Dim linesElem As MSXML2.IXMLDOMElement
    Dim mTextsElem As MSXML2.IXMLDOMElement
    Dim textsElem As MSXML2.IXMLDOMElement
    Dim imagesElem As MSXML2.IXMLDOMElement
    Dim circlesElem As MSXML2.IXMLDOMElement
    Dim attrDefsElem As MSXML2.IXMLDOMElement
    
    Dim originElem As MSXML2.IXMLDOMElement
    Dim layer As IntelliCAD.layer
    
    Dim entity As IntelliCAD.entity
    Dim origin As IntelliCAD.point
    
    Set origin = Block.origin
    Set originElem = ObjDoc.createElement("Origin")
    ObjRoot.appendChild originElem
    originElem.setAttribute "X", origin.x
    originElem.setAttribute "Y", origin.y
    originElem.setAttribute "Z", origin.Z

    
    Set blockInsertsElem = ObjDoc.createElement("BlockInserts")
    ObjRoot.appendChild blockInsertsElem
    
    Set lwPolylinesElem = ObjDoc.createElement("LWPolylines")
    ObjRoot.appendChild lwPolylinesElem
    
    Set polyLinesElem = ObjDoc.createElement("Polylines")
    ObjRoot.appendChild polyLinesElem
    
    Set linesElem = ObjDoc.createElement("Lines")
    ObjRoot.appendChild linesElem
    
    Set mTextsElem = ObjDoc.createElement("MTexts")
    ObjRoot.appendChild mTextsElem
    
    Set textsElem = ObjDoc.createElement("Texts")
    ObjRoot.appendChild textsElem
    
    Set imagesElem = ObjDoc.createElement("Images")
    ObjRoot.appendChild imagesElem
    
    Set circlesElem = ObjDoc.createElement("Circles")
    ObjRoot.appendChild circlesElem
    
    Set attrDefsElem = ObjDoc.createElement("AttributeDefs")
    ObjRoot.appendChild attrDefsElem
    
    For Each entity In Block    '.ModelSpace
        CheckBlockInsert entity, ObjDoc, blockInsertsElem
        CheckLWPolyline entity, ObjDoc, lwPolylinesElem
        CheckPolyline entity, ObjDoc, polyLinesElem
        CheckLine entity, ObjDoc, linesElem
        CheckMText entity, ObjDoc, mTextsElem
        CheckText entity, ObjDoc, textsElem
        CheckImage entity, ObjDoc, imagesElem
        CheckCircle entity, ObjDoc, circlesElem
        CheckAttributeDef entity, ObjDoc, attrDefsElem
    Next entity
    
End Sub
Public Sub CheckBlockInsert(ent As IntelliCAD.entity, ObjDoc As MSXML2.DOMDocument60, ObjRoot As MSXML2.IXMLDOMElement)
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
    
    Dim i As Integer
    
    If TypeOf ent Is IntelliCAD.blockInsert Then
        Set blkInsert = ent
        Set blockInsertElem = ObjDoc.createElement("BlockInsert")
        ObjRoot.appendChild blockInsertElem
        blockInsertElem.setAttribute "Name", blkInsert.Name
        blockInsertElem.setAttribute "Layer", blkInsert.layer
        blockInsertElem.setAttribute "Color", blkInsert.Color.ColorIndex
        blockInsertElem.setAttribute "Rotation", blkInsert.Rotation
        blockInsertElem.setAttribute "Visible", blkInsert.Visible
        blockInsertElem.setAttribute "XScaleFactor", blkInsert.XScaleFactor
        blockInsertElem.setAttribute "YScaleFactor", blkInsert.YScaleFactor
        blockInsertElem.setAttribute "ZScaleFactor", blkInsert.ZScaleFactor
        blockInsertElem.setAttribute "Handle", blkInsert.Handle
        
        Set insertPoint = blkInsert.insertionPoint
        Set insertPointElem = ObjDoc.createElement("InsertionPoint")
        blockInsertElem.appendChild insertPointElem
        insertPointElem.setAttribute "X", insertPoint.x
        insertPointElem.setAttribute "Y", insertPoint.y
        insertPointElem.setAttribute "Z", insertPoint.Z
        
        Set attributesElem = ObjDoc.createElement("Attributes")
        blockInsertElem.appendChild attributesElem
        
        'Loop Through Block Attributes
        If blkInsert.HasAttributes Then
            Set attribs = blkInsert.GetAttributes()
            For i = 0 To attribs.Count - 1
                Set attr = attribs.Item(i)
                Set attribElem = ObjDoc.createElement("Attribute")
                attributesElem.appendChild attribElem
                attribElem.setAttribute "TagString", attr.TagString
                attribElem.setAttribute "TextString", attr.TextString
                attribElem.setAttribute "ScaleFactor", attr.ScaleFactor
                attribElem.setAttribute "Rotation", attr.Rotation
                attribElem.setAttribute "Invisible", attr.Invisible
                attribElem.setAttribute "Height", attr.Height
                
                Set textInsertPoint = attr.TextAlignmentPoint
                Set textInsertPointElem = ObjDoc.createElement("TextAlignmentPoint")
                attribElem.appendChild textInsertPointElem
                textInsertPointElem.setAttribute "X", textInsertPoint.x
                textInsertPointElem.setAttribute "Y", textInsertPoint.y
                textInsertPointElem.setAttribute "Z", textInsertPoint.Z
                
                Set attribInsertPoint = attr.insertionPoint
                Set attribInsertPointElem = ObjDoc.createElement("InsertionPoint")
                attribElem.appendChild attribInsertPointElem
                attribInsertPointElem.setAttribute "X", attribInsertPoint.x
                attribInsertPointElem.setAttribute "Y", attribInsertPoint.y
                attribInsertPointElem.setAttribute "Z", attribInsertPoint.Z
            Next i
        End If
    End If
End Sub
Public Sub CheckLWPolyline(ent As IntelliCAD.entity, ObjDoc As MSXML2.DOMDocument60, ObjRoot As MSXML2.IXMLDOMElement)
    Dim polyLineElem As MSXML2.IXMLDOMElement
    Dim pointElem As MSXML2.IXMLDOMElement
    Dim coordinatesElem As MSXML2.IXMLDOMElement
    Dim widthsElem As IXMLDOMElement
    Dim widthElem As IXMLDOMElement
    
    Dim startWidth As Double
    Dim endWidth As Double
    Dim polyL As IntelliCAD.lwPolyline
    Dim pts As IntelliCAD.points
    Dim pt As IntelliCAD.point
    
    Dim i As Integer
    
    If TypeOf ent Is IntelliCAD.lwPolyline Then
        Set polyL = ent
        Set polyLineElem = ObjDoc.createElement("LWPolyline")
        ObjRoot.appendChild polyLineElem
        polyLineElem.setAttribute "Layer", polyL.layer
        polyLineElem.setAttribute "Color", polyL.Color.ColorIndex
        polyLineElem.setAttribute "Thickness", polyL.Thickness
        polyLineElem.setAttribute "LineType", polyL.LineType
        polyLineElem.setAttribute "LineWeight", polyL.LineWeight
        polyLineElem.setAttribute "Handle", polyL.Handle
        polyLineElem.setAttribute "Closed", polyL.Closed
        polyLineElem.setAttribute "LinetypeScale", polyL.LinetypeScale
        polyLineElem.setAttribute "Length", polyL.Length
        
        Set coordinatesElem = ObjDoc.createElement("Coordinates")
        polyLineElem.appendChild coordinatesElem
            
        For i = 0 To polyL.Coordinates.Count - 1
            Set pt = polyL.Coordinates.Item(i)
            Set pointElem = ObjDoc.createElement("Point")
            coordinatesElem.appendChild pointElem
            pointElem.setAttribute "X", pt.x
            pointElem.setAttribute "Y", pt.y
            pointElem.setAttribute "Z", pt.Z
        Next i
        
        Set widthsElem = ObjDoc.createElement("SegmentWidths")
        polyLineElem.appendChild widthsElem
        
        For i = 0 To polyL.Coordinates.Count - 1
            polyL.GetWidth i, startWidth, endWidth
            Set widthElem = ObjDoc.createElement("Segment")
            widthsElem.appendChild widthElem
            widthElem.setAttribute "StartWidth", startWidth
            widthElem.setAttribute "EndWidth", endWidth
        Next i
    End If

End Sub
Public Sub CheckPolyline(ent As IntelliCAD.entity, ObjDoc As MSXML2.DOMDocument60, ObjRoot As MSXML2.IXMLDOMElement)
    Dim polyLineElem As MSXML2.IXMLDOMElement
    Dim pointElem As MSXML2.IXMLDOMElement
    Dim coordinatesElem As MSXML2.IXMLDOMElement
    
    Dim polyL As IntelliCAD.polyline
    Dim pts As IntelliCAD.points
    Dim pt As IntelliCAD.point
    
    Dim i As Integer
    
    If TypeOf ent Is IntelliCAD.polyline Then
        Set polyL = ent
        Set polyLineElem = ObjDoc.createElement("Polyline")
        ObjRoot.appendChild polyLineElem
        polyLineElem.setAttribute "Layer", polyL.layer
        polyLineElem.setAttribute "Color", polyL.Color.ColorIndex
        polyLineElem.setAttribute "Thickness", polyL.Thickness
        polyLineElem.setAttribute "LineType", polyL.LineType
        polyLineElem.setAttribute "LineWeight", polyL.LineWeight
        polyLineElem.setAttribute "Handle", polyL.Handle
        polyLineElem.setAttribute "Closed", polyL.Closed
        polyLineElem.setAttribute "LinetypeScale", polyL.LinetypeScale
        polyLineElem.setAttribute "Length", polyL.Length
        
        Set coordinatesElem = ObjDoc.createElement("Coordinates")
        polyLineElem.appendChild coordinatesElem
            
        For i = 0 To polyL.Coordinates.Count - 1
            Set pt = polyL.Coordinates.Item(i)
            Set pointElem = ObjDoc.createElement("Point")
            coordinatesElem.appendChild pointElem
            pointElem.setAttribute "X", pt.x
            pointElem.setAttribute "Y", pt.y
            pointElem.setAttribute "Z", pt.Z
        Next i
    End If

End Sub
Public Sub CheckLine(ent As IntelliCAD.entity, ObjDoc As MSXML2.DOMDocument60, ObjRoot As MSXML2.IXMLDOMElement)
    Dim lineElem As MSXML2.IXMLDOMElement
    Dim startPointElem As MSXML2.IXMLDOMElement
    Dim endPointElem As MSXML2.IXMLDOMElement
    Dim coordinatesElem As MSXML2.IXMLDOMElement
    
    Dim lin As IntelliCAD.line
    Dim stPoint As IntelliCAD.point
    Dim enPoint As IntelliCAD.point
    
    If TypeOf ent Is IntelliCAD.line Then
        Set lin = ent
        Set lineElem = ObjDoc.createElement("Line")
        ObjRoot.appendChild lineElem
        lineElem.setAttribute "Layer", lin.layer
        lineElem.setAttribute "Color", lin.Color.ColorIndex
        lineElem.setAttribute "Thickness", lin.Thickness
        lineElem.setAttribute "LineType", lin.LineType
        lineElem.setAttribute "LineWeight", lin.LineWeight
        lineElem.setAttribute "Handle", lin.Handle
        lineElem.setAttribute "Visible", lin.Visible
        
        Set stPoint = lin.startPoint
        Set startPointElem = ObjDoc.createElement("StartPoint")
        lineElem.appendChild startPointElem
        startPointElem.setAttribute "X", stPoint.x
        startPointElem.setAttribute "Y", stPoint.y
        startPointElem.setAttribute "Z", stPoint.Z
        
        Set enPoint = lin.endPoint
        Set endPointElem = ObjDoc.createElement("EndPoint")
        lineElem.appendChild endPointElem
        endPointElem.setAttribute "X", enPoint.x
        endPointElem.setAttribute "Y", enPoint.y
        endPointElem.setAttribute "Z", enPoint.Z
    End If
    
End Sub

Public Sub CheckMText(ent As IntelliCAD.entity, ObjDoc As MSXML2.DOMDocument60, ObjRoot As MSXML2.IXMLDOMElement)
    Dim textElem As MSXML2.IXMLDOMElement
    Dim insertPointElem As MSXML2.IXMLDOMElement
    
    Dim text As IntelliCAD.mText
    Dim insertPoint As IntelliCAD.point
    
    If TypeOf ent Is IntelliCAD.mText Then
        Set text = ent
        Set textElem = ObjDoc.createElement("MText")
        ObjRoot.appendChild textElem
        textElem.setAttribute "Layer", text.layer
        textElem.setAttribute "Color", text.Color.ColorIndex
        textElem.setAttribute "TextString", text.TextString
        textElem.setAttribute "Handle", text.Handle
        textElem.setAttribute "Width", text.Width
        textElem.setAttribute "Height", text.Height
        textElem.setAttribute "Rotation", text.Rotation
        textElem.setAttribute "StyleName", text.styleName
          
        Set insertPoint = text.insertionPoint
        Set insertPointElem = ObjDoc.createElement("InsertionPoint")
        textElem.appendChild insertPointElem
        insertPointElem.setAttribute "X", insertPoint.x
        insertPointElem.setAttribute "Y", insertPoint.y
        insertPointElem.setAttribute "Z", insertPoint.Z
    End If

End Sub

Public Sub CheckText(ent As IntelliCAD.entity, ObjDoc As MSXML2.DOMDocument60, ObjRoot As MSXML2.IXMLDOMElement)
    Dim textElem As MSXML2.IXMLDOMElement
    Dim insertPointElem As MSXML2.IXMLDOMElement
    
    Dim text As IntelliCAD.text
    Dim insertPoint As IntelliCAD.point
    
    If TypeOf ent Is IntelliCAD.text Then
        Set text = ent
        Set textElem = ObjDoc.createElement("Text")
        ObjRoot.appendChild textElem
        textElem.setAttribute "Layer", text.layer
        textElem.setAttribute "Color", text.Color.ColorIndex
        textElem.setAttribute "TextString", text.TextString
        textElem.setAttribute "Handle", text.Handle
        textElem.setAttribute "Height", text.Height
        textElem.setAttribute "Rotation", text.Rotation
        textElem.setAttribute "ScaleFactor", text.ScaleFactor
        textElem.setAttribute "StyleName", text.styleName
        textElem.setAttribute "HorizontalAlignment", text.HorizontalAlignment
        textElem.setAttribute "VerticalAlignment", text.VerticalAlignment
          
        Set insertPoint = text.insertionPoint
        Set insertPointElem = ObjDoc.createElement("InsertionPoint")
        textElem.appendChild insertPointElem
        insertPointElem.setAttribute "X", insertPoint.x
        insertPointElem.setAttribute "Y", insertPoint.y
        insertPointElem.setAttribute "Z", insertPoint.Z
        
        Set insertPoint = text.TextAlignmentPoint
        Set insertPointElem = ObjDoc.createElement("TextAlignmentPoint")
        textElem.appendChild insertPointElem
        insertPointElem.setAttribute "X", insertPoint.x
        insertPointElem.setAttribute "Y", insertPoint.y
        insertPointElem.setAttribute "Z", insertPoint.Z
    End If

End Sub

Public Sub CheckImage(ent As IntelliCAD.entity, ObjDoc As MSXML2.DOMDocument60, ObjRoot As MSXML2.IXMLDOMElement)
    Dim imgElem As MSXML2.IXMLDOMElement
    Dim insertPointElem As MSXML2.IXMLDOMElement
    
    Dim img As IntelliCAD.Image
    Dim insertPoint As IntelliCAD.point
    
    If TypeOf ent Is IntelliCAD.Image Then
        Set img = ent
        Set imgElem = ObjDoc.createElement("Image")
        ObjRoot.appendChild imgElem
        imgElem.setAttribute "ImageFile", img.ImageFile
        imgElem.setAttribute "Layer", img.layer
        imgElem.setAttribute "Color", img.Color.ColorIndex
        imgElem.setAttribute "Width", img.Width
        imgElem.setAttribute "Height", img.Height
        imgElem.setAttribute "ImageWidth", img.ImageWidth
        imgElem.setAttribute "ImageHeight", img.ImageHeight
        imgElem.setAttribute "Visible", img.Visible
        imgElem.setAttribute "Handle", img.Handle
          
        Set insertPoint = img.origin
        Set insertPointElem = ObjDoc.createElement("Origin")
        imgElem.appendChild insertPointElem
        insertPointElem.setAttribute "X", insertPoint.x
        insertPointElem.setAttribute "Y", insertPoint.y
        insertPointElem.setAttribute "Z", insertPoint.Z
    End If

End Sub

Public Sub CheckTextStyles(txtSty As IntelliCAD.textStyle, ObjDoc As MSXML2.DOMDocument60, ObjRoot As MSXML2.IXMLDOMElement)
    Dim textStyleElem As MSXML2.IXMLDOMElement
    
    Set textStyleElem = ObjDoc.createElement("TextStyle")
    ObjRoot.appendChild textStyleElem
    
    textStyleElem.setAttribute "Name", txtSty.Name
    textStyleElem.setAttribute "Width", txtSty.Width
    textStyleElem.setAttribute "Height", txtSty.Height
    textStyleElem.setAttribute "FontFile", txtSty.FontFile
    textStyleElem.setAttribute "BigFontFile", txtSty.BigFontFile
    
End Sub
Public Sub CheckLayers(layer As IntelliCAD.layer, ObjDoc As MSXML2.DOMDocument60, ObjRoot As MSXML2.IXMLDOMElement)
    Dim layerElem As MSXML2.IXMLDOMElement
    Dim insertPointElem As MSXML2.IXMLDOMElement
    
    Set layerElem = ObjDoc.createElement("Layer")
    ObjRoot.appendChild layerElem
    layerElem.setAttribute "Name", layer.Name
    layerElem.setAttribute "ColorIndex", layer.Color.ColorIndex
    layerElem.setAttribute "ColorMethod", layer.Color.ColorMethod
    layerElem.setAttribute "EntityColor", layer.Color.EntityColor
    layerElem.setAttribute "ColorRed", layer.Color.Red
    layerElem.setAttribute "ColorBlue", layer.Color.Blue
    layerElem.setAttribute "ColorGreen", layer.Color.Green
    layerElem.setAttribute "Freeze", layer.Freeze
    layerElem.setAttribute "Lock", layer.Lock
    layerElem.setAttribute "LineType", layer.LineType
    layerElem.setAttribute "LineWeight", layer.LineWeight
    layerElem.setAttribute "Handle", layer.Handle
    layerElem.setAttribute "LayerOn", layer.LayerOn
    layerElem.setAttribute "PlotStyleName", layer.PlotStyleName
    layerElem.setAttribute "Plottable", layer.Plottable
    
End Sub
Public Sub CheckCircle(ent As IntelliCAD.entity, ObjDoc As MSXML2.DOMDocument60, ObjRoot As MSXML2.IXMLDOMElement)
    Dim circleElem As MSXML2.IXMLDOMElement
    Dim centerPointElem As MSXML2.IXMLDOMElement
    
    Dim circ As IntelliCAD.Circle
    Dim centerPoint As IntelliCAD.point
    
    If TypeOf ent Is IntelliCAD.Circle Then
        Set circ = ent
        Set circleElem = ObjDoc.createElement("Circle")
        ObjRoot.appendChild circleElem
        circleElem.setAttribute "Layer", circ.layer
        circleElem.setAttribute "Color", circ.Color.ColorIndex
        circleElem.setAttribute "Radius", circ.Radius
        circleElem.setAttribute "Handle", circ.Handle
          
        Set centerPoint = circ.center
        Set centerPointElem = ObjDoc.createElement("Center")
        circleElem.appendChild centerPointElem
        centerPointElem.setAttribute "X", centerPoint.x
        centerPointElem.setAttribute "Y", centerPoint.y
        centerPointElem.setAttribute "Z", centerPoint.Z
    End If

End Sub
Public Sub CheckAttributeDef(ent As IntelliCAD.entity, ObjDoc As MSXML2.DOMDocument60, ObjRoot As MSXML2.IXMLDOMElement)
    Dim attrDefElem As MSXML2.IXMLDOMElement
    Dim pointElem As MSXML2.IXMLDOMElement
    
    Dim attrDef As IntelliCAD.AttributeDef
    Dim pnt As IntelliCAD.point
    
    If TypeOf ent Is IntelliCAD.AttributeDef Then
        Set attrDef = ent
        Set attrDefElem = ObjDoc.createElement("AttributeDef")
        ObjRoot.appendChild attrDefElem
        attrDefElem.setAttribute "Layer", attrDef.layer
        attrDefElem.setAttribute "Color", attrDef.Color.ColorIndex
        attrDefElem.setAttribute "Handle", attrDef.Handle
        attrDefElem.setAttribute "TextString", attrDef.TextString
        attrDefElem.setAttribute "TagString", attrDef.TagString
        attrDefElem.setAttribute "Height", attrDef.Height
        attrDefElem.setAttribute "Mode", attrDef.Mode
        attrDefElem.setAttribute "Visible", attrDef.Visible
        attrDefElem.setAttribute "Invisible", attrDef.Invisible
        attrDefElem.setAttribute "StyleName", attrDef.styleName
        attrDefElem.setAttribute "Rotation", attrDef.Rotation
          
        Set pnt = attrDef.insertionPoint
        Set pointElem = ObjDoc.createElement("InsertionPoint")
        attrDefElem.appendChild pointElem
        pointElem.setAttribute "X", pnt.x
        pointElem.setAttribute "Y", pnt.y
        pointElem.setAttribute "Z", pnt.Z
        
        Set pnt = attrDef.TextAlignmentPoint
        Set pointElem = ObjDoc.createElement("TextAlignmentPoint")
        attrDefElem.appendChild pointElem
        pointElem.setAttribute "X", pnt.x
        pointElem.setAttribute "Y", pnt.y
        pointElem.setAttribute "Z", pnt.Z
    End If

End Sub


