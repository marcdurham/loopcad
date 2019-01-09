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
    Dim objDoc As MSXML2.DOMDocument60
    Dim objNode As MSXML2.IXMLDOMNode
    Dim objRoot As MSXML2.IXMLDOMElement
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
    Dim block As IntelliCAD.block
    Dim entity As IntelliCAD.entity
    
    If IsMissing(outputFilePath) Then
        outputFilePath = ShowSave("Save Scan File As...", "CadDrawingScan", "Cad Drawing Scan (*.cds)", "*.cds", ActiveDocument.path)
        If outputFilePath = "" Then
            Exit Sub
        End If
    End If
    
    Set objDoc = New DOMDocument60
    objDoc.resolveExternals = True
    
    Set objNode = objDoc.createProcessingInstruction("xml", "version='1.0' encoding='UTF-8'")
    Set objNode = objDoc.insertBefore(objNode, objDoc.childNodes.Item(0))
    Set objRoot = objDoc.createElement("CadData")
    Set objDoc.documentElement = objRoot
    
    Set documentElem = objDoc.createElement("Document")
    objRoot.appendChild documentElem
    documentElem.setAttribute "FullName", ActiveDocument.FullName
    documentElem.setAttribute "Name", ActiveDocument.name
    documentElem.setAttribute "Path", ActiveDocument.path
    
    Set layersElem = objDoc.createElement("Layers")
    objRoot.appendChild layersElem
    
    For Each layer In ActiveDocument.Layers
        CheckLayers layer, objDoc, layersElem
    Next layer
    
    Dim textStylesElem As MSXML2.IXMLDOMElement
    Dim txtSty As IntelliCAD.textStyle
    Set textStylesElem = objDoc.createElement("TextStyles")
    objRoot.appendChild textStylesElem
    
    For Each txtSty In ActiveDocument.TextStyles
        CheckTextStyles txtSty, objDoc, textStylesElem
    Next txtSty
    
    Set blocksElem = objDoc.createElement("Blocks")
    objRoot.appendChild blocksElem
    
    For Each block In ActiveDocument.Blocks
        Set blockElem = objDoc.createElement("Block")
        blocksElem.appendChild blockElem
        blockElem.setAttribute "Name", block.name
        blockElem.setAttribute "IsLayout", block.IsLayout
        blockElem.setAttribute "IsXRef", block.IsXRef
        
        ScanBlock block, objDoc, blockElem
    Next block
    
    objDoc.Save outputFilePath

End Sub
Public Sub ScanBlock(block As IntelliCAD.block, objDoc As MSXML2.DOMDocument60, objRoot As MSXML2.IXMLDOMElement)
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
    
    Set origin = block.origin
    Set originElem = objDoc.createElement("Origin")
    objRoot.appendChild originElem
    originElem.setAttribute "X", origin.x
    originElem.setAttribute "Y", origin.y
    originElem.setAttribute "Z", origin.Z

    
    Set blockInsertsElem = objDoc.createElement("BlockInserts")
    objRoot.appendChild blockInsertsElem
    
    Set lwPolylinesElem = objDoc.createElement("LWPolylines")
    objRoot.appendChild lwPolylinesElem
    
    Set polyLinesElem = objDoc.createElement("Polylines")
    objRoot.appendChild polyLinesElem
    
    Set linesElem = objDoc.createElement("Lines")
    objRoot.appendChild linesElem
    
    Set mTextsElem = objDoc.createElement("MTexts")
    objRoot.appendChild mTextsElem
    
    Set textsElem = objDoc.createElement("Texts")
    objRoot.appendChild textsElem
    
    Set imagesElem = objDoc.createElement("Images")
    objRoot.appendChild imagesElem
    
    Set circlesElem = objDoc.createElement("Circles")
    objRoot.appendChild circlesElem
    
    Set attrDefsElem = objDoc.createElement("AttributeDefs")
    objRoot.appendChild attrDefsElem
    
    For Each entity In block    '.ModelSpace
        CheckBlockInsert entity, objDoc, blockInsertsElem
        CheckLWPolyline entity, objDoc, lwPolylinesElem
        CheckPolyline entity, objDoc, polyLinesElem
        CheckLine entity, objDoc, linesElem
        CheckMText entity, objDoc, mTextsElem
        CheckText entity, objDoc, textsElem
        CheckImage entity, objDoc, imagesElem
        CheckCircle entity, objDoc, circlesElem
        CheckAttributeDef entity, objDoc, attrDefsElem
    Next entity
    
End Sub
Public Sub CheckBlockInsert(ent As IntelliCAD.entity, objDoc As MSXML2.DOMDocument60, objRoot As MSXML2.IXMLDOMElement)
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
        Set blockInsertElem = objDoc.createElement("BlockInsert")
        objRoot.appendChild blockInsertElem
        blockInsertElem.setAttribute "Name", blkInsert.name
        blockInsertElem.setAttribute "Layer", blkInsert.layer
        blockInsertElem.setAttribute "Color", blkInsert.Color.ColorIndex
        blockInsertElem.setAttribute "Rotation", blkInsert.Rotation
        blockInsertElem.setAttribute "Visible", blkInsert.Visible
        blockInsertElem.setAttribute "XScaleFactor", blkInsert.XScaleFactor
        blockInsertElem.setAttribute "YScaleFactor", blkInsert.YScaleFactor
        blockInsertElem.setAttribute "ZScaleFactor", blkInsert.ZScaleFactor
        blockInsertElem.setAttribute "Handle", blkInsert.Handle
        
        Set insertPoint = blkInsert.insertionPoint
        Set insertPointElem = objDoc.createElement("InsertionPoint")
        blockInsertElem.appendChild insertPointElem
        insertPointElem.setAttribute "X", insertPoint.x
        insertPointElem.setAttribute "Y", insertPoint.y
        insertPointElem.setAttribute "Z", insertPoint.Z
        
        Set attributesElem = objDoc.createElement("Attributes")
        blockInsertElem.appendChild attributesElem
        
        'Loop Through Block Attributes
        If blkInsert.HasAttributes Then
            Set attribs = blkInsert.GetAttributes()
            For i = 0 To attribs.Count - 1
                Set attr = attribs.Item(i)
                Set attribElem = objDoc.createElement("Attribute")
                attributesElem.appendChild attribElem
                attribElem.setAttribute "TagString", attr.TagString
                attribElem.setAttribute "TextString", attr.TextString
                attribElem.setAttribute "ScaleFactor", attr.ScaleFactor
                attribElem.setAttribute "Rotation", attr.Rotation
                attribElem.setAttribute "Invisible", attr.Invisible
                attribElem.setAttribute "Height", attr.Height
                
                Set textInsertPoint = attr.TextAlignmentPoint
                Set textInsertPointElem = objDoc.createElement("TextAlignmentPoint")
                attribElem.appendChild textInsertPointElem
                textInsertPointElem.setAttribute "X", textInsertPoint.x
                textInsertPointElem.setAttribute "Y", textInsertPoint.y
                textInsertPointElem.setAttribute "Z", textInsertPoint.Z
                
                Set attribInsertPoint = attr.insertionPoint
                Set attribInsertPointElem = objDoc.createElement("InsertionPoint")
                attribElem.appendChild attribInsertPointElem
                attribInsertPointElem.setAttribute "X", attribInsertPoint.x
                attribInsertPointElem.setAttribute "Y", attribInsertPoint.y
                attribInsertPointElem.setAttribute "Z", attribInsertPoint.Z
            Next i
        End If
    End If
End Sub
Public Sub CheckLWPolyline(ent As IntelliCAD.entity, objDoc As MSXML2.DOMDocument60, objRoot As MSXML2.IXMLDOMElement)
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
        Set polyLineElem = objDoc.createElement("LWPolyline")
        objRoot.appendChild polyLineElem
        polyLineElem.setAttribute "Layer", polyL.layer
        polyLineElem.setAttribute "Color", polyL.Color.ColorIndex
        polyLineElem.setAttribute "Thickness", polyL.Thickness
        polyLineElem.setAttribute "LineType", polyL.LineType
        polyLineElem.setAttribute "LineWeight", polyL.LineWeight
        polyLineElem.setAttribute "Handle", polyL.Handle
        polyLineElem.setAttribute "Closed", polyL.Closed
        polyLineElem.setAttribute "LinetypeScale", polyL.LinetypeScale
        polyLineElem.setAttribute "Length", polyL.Length
        
        Set coordinatesElem = objDoc.createElement("Coordinates")
        polyLineElem.appendChild coordinatesElem
            
        For i = 0 To polyL.Coordinates.Count - 1
            Set pt = polyL.Coordinates.Item(i)
            Set pointElem = objDoc.createElement("Point")
            coordinatesElem.appendChild pointElem
            pointElem.setAttribute "X", pt.x
            pointElem.setAttribute "Y", pt.y
            pointElem.setAttribute "Z", pt.Z
        Next i
        
        Set widthsElem = objDoc.createElement("SegmentWidths")
        polyLineElem.appendChild widthsElem
        
        For i = 0 To polyL.Coordinates.Count - 1
            polyL.GetWidth i, startWidth, endWidth
            Set widthElem = objDoc.createElement("Segment")
            widthsElem.appendChild widthElem
            widthElem.setAttribute "StartWidth", startWidth
            widthElem.setAttribute "EndWidth", endWidth
        Next i
    End If

End Sub
Public Sub CheckPolyline(ent As IntelliCAD.entity, objDoc As MSXML2.DOMDocument60, objRoot As MSXML2.IXMLDOMElement)
    Dim polyLineElem As MSXML2.IXMLDOMElement
    Dim pointElem As MSXML2.IXMLDOMElement
    Dim coordinatesElem As MSXML2.IXMLDOMElement
    
    Dim polyL As IntelliCAD.polyline
    Dim pts As IntelliCAD.points
    Dim pt As IntelliCAD.point
    
    Dim i As Integer
    
    If TypeOf ent Is IntelliCAD.polyline Then
        Set polyL = ent
        Set polyLineElem = objDoc.createElement("Polyline")
        objRoot.appendChild polyLineElem
        polyLineElem.setAttribute "Layer", polyL.layer
        polyLineElem.setAttribute "Color", polyL.Color.ColorIndex
        polyLineElem.setAttribute "Thickness", polyL.Thickness
        polyLineElem.setAttribute "LineType", polyL.LineType
        polyLineElem.setAttribute "LineWeight", polyL.LineWeight
        polyLineElem.setAttribute "Handle", polyL.Handle
        polyLineElem.setAttribute "Closed", polyL.Closed
        polyLineElem.setAttribute "LinetypeScale", polyL.LinetypeScale
        polyLineElem.setAttribute "Length", polyL.Length
        
        Set coordinatesElem = objDoc.createElement("Coordinates")
        polyLineElem.appendChild coordinatesElem
            
        For i = 0 To polyL.Coordinates.Count - 1
            Set pt = polyL.Coordinates.Item(i)
            Set pointElem = objDoc.createElement("Point")
            coordinatesElem.appendChild pointElem
            pointElem.setAttribute "X", pt.x
            pointElem.setAttribute "Y", pt.y
            pointElem.setAttribute "Z", pt.Z
        Next i
    End If

End Sub
Public Sub CheckLine(ent As IntelliCAD.entity, objDoc As MSXML2.DOMDocument60, objRoot As MSXML2.IXMLDOMElement)
    Dim lineElem As MSXML2.IXMLDOMElement
    Dim startPointElem As MSXML2.IXMLDOMElement
    Dim endPointElem As MSXML2.IXMLDOMElement
    Dim coordinatesElem As MSXML2.IXMLDOMElement
    
    Dim lin As IntelliCAD.line
    Dim stPoint As IntelliCAD.point
    Dim enPoint As IntelliCAD.point
    
    If TypeOf ent Is IntelliCAD.line Then
        Set lin = ent
        Set lineElem = objDoc.createElement("Line")
        objRoot.appendChild lineElem
        lineElem.setAttribute "Layer", lin.layer
        lineElem.setAttribute "Color", lin.Color.ColorIndex
        lineElem.setAttribute "Thickness", lin.Thickness
        lineElem.setAttribute "LineType", lin.LineType
        lineElem.setAttribute "LineWeight", lin.LineWeight
        lineElem.setAttribute "Handle", lin.Handle
        lineElem.setAttribute "Visible", lin.Visible
        
        Set stPoint = lin.startPoint
        Set startPointElem = objDoc.createElement("StartPoint")
        lineElem.appendChild startPointElem
        startPointElem.setAttribute "X", stPoint.x
        startPointElem.setAttribute "Y", stPoint.y
        startPointElem.setAttribute "Z", stPoint.Z
        
        Set enPoint = lin.endPoint
        Set endPointElem = objDoc.createElement("EndPoint")
        lineElem.appendChild endPointElem
        endPointElem.setAttribute "X", enPoint.x
        endPointElem.setAttribute "Y", enPoint.y
        endPointElem.setAttribute "Z", enPoint.Z
    End If
    
End Sub

Public Sub CheckMText(ent As IntelliCAD.entity, objDoc As MSXML2.DOMDocument60, objRoot As MSXML2.IXMLDOMElement)
    Dim textElem As MSXML2.IXMLDOMElement
    Dim insertPointElem As MSXML2.IXMLDOMElement
    
    Dim text As IntelliCAD.mText
    Dim insertPoint As IntelliCAD.point
    
    If TypeOf ent Is IntelliCAD.mText Then
        Set text = ent
        Set textElem = objDoc.createElement("MText")
        objRoot.appendChild textElem
        textElem.setAttribute "Layer", text.layer
        textElem.setAttribute "Color", text.Color.ColorIndex
        textElem.setAttribute "TextString", text.TextString
        textElem.setAttribute "Handle", text.Handle
        textElem.setAttribute "Width", text.Width
        textElem.setAttribute "Height", text.Height
        textElem.setAttribute "Rotation", text.Rotation
        textElem.setAttribute "StyleName", text.styleName
          
        Set insertPoint = text.insertionPoint
        Set insertPointElem = objDoc.createElement("InsertionPoint")
        textElem.appendChild insertPointElem
        insertPointElem.setAttribute "X", insertPoint.x
        insertPointElem.setAttribute "Y", insertPoint.y
        insertPointElem.setAttribute "Z", insertPoint.Z
    End If

End Sub

Public Sub CheckText(ent As IntelliCAD.entity, objDoc As MSXML2.DOMDocument60, objRoot As MSXML2.IXMLDOMElement)
    Dim textElem As MSXML2.IXMLDOMElement
    Dim insertPointElem As MSXML2.IXMLDOMElement
    
    Dim text As IntelliCAD.text
    Dim insertPoint As IntelliCAD.point
    
    If TypeOf ent Is IntelliCAD.text Then
        Set text = ent
        Set textElem = objDoc.createElement("Text")
        objRoot.appendChild textElem
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
        Set insertPointElem = objDoc.createElement("InsertionPoint")
        textElem.appendChild insertPointElem
        insertPointElem.setAttribute "X", insertPoint.x
        insertPointElem.setAttribute "Y", insertPoint.y
        insertPointElem.setAttribute "Z", insertPoint.Z
        
        Set insertPoint = text.TextAlignmentPoint
        Set insertPointElem = objDoc.createElement("TextAlignmentPoint")
        textElem.appendChild insertPointElem
        insertPointElem.setAttribute "X", insertPoint.x
        insertPointElem.setAttribute "Y", insertPoint.y
        insertPointElem.setAttribute "Z", insertPoint.Z
    End If

End Sub

Public Sub CheckImage(ent As IntelliCAD.entity, objDoc As MSXML2.DOMDocument60, objRoot As MSXML2.IXMLDOMElement)
    Dim imgElem As MSXML2.IXMLDOMElement
    Dim insertPointElem As MSXML2.IXMLDOMElement
    
    Dim img As IntelliCAD.Image
    Dim insertPoint As IntelliCAD.point
    
    If TypeOf ent Is IntelliCAD.Image Then
        Set img = ent
        Set imgElem = objDoc.createElement("Image")
        objRoot.appendChild imgElem
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
        Set insertPointElem = objDoc.createElement("Origin")
        imgElem.appendChild insertPointElem
        insertPointElem.setAttribute "X", insertPoint.x
        insertPointElem.setAttribute "Y", insertPoint.y
        insertPointElem.setAttribute "Z", insertPoint.Z
    End If

End Sub

Public Sub CheckTextStyles(txtSty As IntelliCAD.textStyle, objDoc As MSXML2.DOMDocument60, objRoot As MSXML2.IXMLDOMElement)
    Dim textStyleElem As MSXML2.IXMLDOMElement
    
    Set textStyleElem = objDoc.createElement("TextStyle")
    objRoot.appendChild textStyleElem
    
    textStyleElem.setAttribute "Name", txtSty.name
    textStyleElem.setAttribute "Width", txtSty.Width
    textStyleElem.setAttribute "Height", txtSty.Height
    textStyleElem.setAttribute "FontFile", txtSty.FontFile
    textStyleElem.setAttribute "BigFontFile", txtSty.BigFontFile
    
End Sub
Public Sub CheckLayers(layer As IntelliCAD.layer, objDoc As MSXML2.DOMDocument60, objRoot As MSXML2.IXMLDOMElement)
    Dim layerElem As MSXML2.IXMLDOMElement
    Dim insertPointElem As MSXML2.IXMLDOMElement
    
    Set layerElem = objDoc.createElement("Layer")
    objRoot.appendChild layerElem
    layerElem.setAttribute "Name", layer.name
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
Public Sub CheckCircle(ent As IntelliCAD.entity, objDoc As MSXML2.DOMDocument60, objRoot As MSXML2.IXMLDOMElement)
    Dim circleElem As MSXML2.IXMLDOMElement
    Dim centerPointElem As MSXML2.IXMLDOMElement
    
    Dim circ As IntelliCAD.Circle
    Dim centerPoint As IntelliCAD.point
    
    If TypeOf ent Is IntelliCAD.Circle Then
        Set circ = ent
        Set circleElem = objDoc.createElement("Circle")
        objRoot.appendChild circleElem
        circleElem.setAttribute "Layer", circ.layer
        circleElem.setAttribute "Color", circ.Color.ColorIndex
        circleElem.setAttribute "Radius", circ.Radius
        circleElem.setAttribute "Handle", circ.Handle
          
        Set centerPoint = circ.center
        Set centerPointElem = objDoc.createElement("Center")
        circleElem.appendChild centerPointElem
        centerPointElem.setAttribute "X", centerPoint.x
        centerPointElem.setAttribute "Y", centerPoint.y
        centerPointElem.setAttribute "Z", centerPoint.Z
    End If

End Sub
Public Sub CheckAttributeDef(ent As IntelliCAD.entity, objDoc As MSXML2.DOMDocument60, objRoot As MSXML2.IXMLDOMElement)
    Dim attrDefElem As MSXML2.IXMLDOMElement
    Dim pointElem As MSXML2.IXMLDOMElement
    
    Dim attrDef As IntelliCAD.AttributeDef
    Dim pnt As IntelliCAD.point
    
    If TypeOf ent Is IntelliCAD.AttributeDef Then
        Set attrDef = ent
        Set attrDefElem = objDoc.createElement("AttributeDef")
        objRoot.appendChild attrDefElem
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
        Set pointElem = objDoc.createElement("InsertionPoint")
        attrDefElem.appendChild pointElem
        pointElem.setAttribute "X", pnt.x
        pointElem.setAttribute "Y", pnt.y
        pointElem.setAttribute "Z", pnt.Z
        
        Set pnt = attrDef.TextAlignmentPoint
        Set pointElem = objDoc.createElement("TextAlignmentPoint")
        attrDefElem.appendChild pointElem
        pointElem.setAttribute "X", pnt.x
        pointElem.setAttribute "Y", pnt.y
        pointElem.setAttribute "Z", pnt.Z
    End If

End Sub


