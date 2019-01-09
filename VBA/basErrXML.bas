Attribute VB_Name = "basErrXML"
Option Explicit
Public Function FSDXNodes2(nodes As Collection) As String
Dim xNode As node
Dim manifoldDone As Boolean
Dim strXMLDoc As String
Dim strNodeRow As String
    Dim objDom As DOMActiveDocument
    Dim objRootElem As IXMLDOMElement
    Dim objMemberElem As IXMLDOMElement
    Dim objMemberRel As IXMLDOMAttribute
    Dim objMemberName As IXMLDOMElement

        strXMLDoc = ""
        
        Set objDom = New DOMActiveDocument
   'objdom.loadXML(
    ' Creates root element
    Set objRootElem = objDom.createElement("nodes")
    objDom.appendChild objRootElem
    

    
    
    
    ' Saves XML data to disk.
    'objDom.Save ("d:\temp\andrew.xml")
    Debug.Print "Text:" & objDom.xml

        For Each xNode In nodes
            ' Creates Member element
            Set objMemberElem = objDom.createElement("node")
            objRootElem.appendChild objMemberElem
            
            
            ' Create elements under "node" element, and
            Set objMemberName = objDom.createElement("name")
            objMemberElem.appendChild objMemberName
            objMemberName.text = xNode.name
            
            
            Dim eleNodeNo As IXMLDOMElement
            Set eleNodeNo = objDom.createElement("name")
            objMemberElem.appendChild eleNodeNo
            eleNodeNo.text = xNode.NodeNumber
            
            
            
            
            

            'Temporary section, omits extra manifold ports
            If xNode.NodeTypeID = NODE_TYPE_MANIFOLD_PORT Then
                If manifoldDone Then
                    GoTo NextxNode
                Else
                    manifoldDone = True
                End If
            End If
            
            
            
            strNodeRow = strNodeRow & vbTab & vbTab & vbTab & vbTab & vbTab & "<name>" & xNode.name & "</name>" & vbCrLf
            strNodeRow = strNodeRow & vbTab & vbTab & vbTab & vbTab & vbTab & "<node_no>" & xNode.NodeNumber & "</node_no>" & vbCrLf
            strNodeRow = strNodeRow & vbTab & vbTab & vbTab & vbTab & vbTab & "<insertion_point><x>" & xNode.insertionPoint.x & "</x><y>" & xNode.insertionPoint.y & "</y><z>" & xNode.insertionPoint.Z & "</z></insertion_point>" & vbCrLf
            strNodeRow = strNodeRow & vbTab & vbTab & vbTab & vbTab & vbTab & "<elevation>" & xNode.elevation & "</elevation>" & vbCrLf
            strNodeRow = strNodeRow & vbTab & vbTab & vbTab & vbTab & vbTab & "<model_number>" & xNode.model & "</model_number>" & vbCrLf
            strNodeRow = strNodeRow & vbTab & vbTab & vbTab & vbTab & vbTab & "<node_type_id>" & xNode.NodeTypeID & "</node_type_id>" & vbCrLf
            strNodeRow = strNodeRow & vbTab & vbTab & vbTab & vbTab & "</node>" & vbCrLf
            strXMLDoc = strXMLDoc & strNodeRow
NextxNode:
        Next xNode
        
        ''no''strXMLDoc = strXMLDoc & vbTab & vbTab & "</nodes>" & vbCrLf
        
        FSDXNodes = strXMLDoc

    
End Function
Public Sub SaveErrorXML()
'Private Sub Create_XML()
   
   Dim objDom As DOMActiveDocument
   Dim objRootElem As IXMLDOMElement
   Dim objMemberElem As IXMLDOMElement
   Dim objMemberRel As IXMLDOMAttribute
   Dim objMemberName As IXMLDOMElement
   
   Set objDom = New DOMActiveDocument
   
   ' Creates root element
   Set objRootElem = objDom.createElement("fsdx_file")
   objDom.appendChild objRootElem
   
   ' Creates Member element
   Set objMemberElem = objDom.createElement("Member")
   objRootElem.appendChild objMemberElem
   
   ' Creates Attribute to the Member Element
   Set objMemberRel = objDom.createAttribute("Relationship")
   objMemberRel.nodeValue = "Father"
   objMemberElem.setAttributeNode objMemberRel
   
   ' Create element under Member element, and
   ' gives value "some guy"
   Set objMemberName = objDom.createElement("Name")
   objMemberElem.appendChild objMemberName
   objMemberName.text = "Some Guy"

    strXSL = "<?xml version=""1.0"" encoding=""UTF-8""?>"
    strXSL = strXSL & "<xsl:stylesheet version=""1.0"" xmlns:xsl=""http://www.w3.org/1999/XSL/Transform"">"
    strXSL = strXSL & "<xsl:output method=""xml"" indent=""yes"" encoding=""UTF-8"" />"
    strXSL = strXSL & "<xsl:template match=""@* | node()"">"
    strXSL = strXSL & "<xsl:copy>"
    strXSL = strXSL & "<xsl:apply-templates select=""@* | node()"" />"
    strXSL = strXSL & "</xsl:copy>"
    strXSL = strXSL & "</xsl:template>"
    strXSL = strXSL & "</xsl:stylesheet>"
    
    Dim xslIndenter As DOMActiveDocument
    Dim indentedXML As DOMActiveDocument
    Set indentedXML = New DOMActiveDocument
    Set xslIndenter = New DOMActiveDocument
    xslIndenter.async = False
    If Not xslIndenter.loadXML(strXSL) Then Debug.Print "Cannot LoadXML!"
    objDom.transformNodeToObject xslIndenter, indentedXML
   ' Saves XML data to disk.
   objDom.Save ("c:\errors.xml")
   Debug.Print "Text:" & objDom.xml
   Debug.Print "Indented Text:" & indentedXML.xml
'End Sub
End Sub

Public Sub ScanCADObjects()
    Dim entit As IntelliCAD.entity
    
    Dim entPoly As AutoCAD.AcadObject
    Dim entPolylines As New Collection
    
    Dim entLWPolyline As IntelliCAD.lwPolyline
    Dim entLWPolylines As New Collection
    
    Dim cadTexts As New Collection
    Dim entText As IntelliCAD.text
    
    Dim cadMTexts As New Collection
    Dim entMText As IntelliCAD.mText
    
    Dim cadBlocks As New Collection
    Dim entBlock As IntelliCAD.blockInsert
    
    Dim previousLayer As String
    previousLayer = ""

    'See which Boxes exist
    For Each entit In ActiveDocument.ModelSpace

        If IsOurLayer(entit.layer) Then
            If previousLayer <> entit.layer Then
                Debug.Print "ScanCADObjects = "; entit.layer
                previousLayer = entit.layer
                ''Debug.Print "   Elevation Box Layer! Entity.EntityName = " & entit.EntityName
            End If
           'Debug.Print "Elevation Box Layer! Entity.EntityName = " & entit.EntityName
            If UCase(entit.EntityName) = UCase("BlockInsert") Then
                Set entBlock = entit
                cadBlocks.Add entBlock
            ElseIf UCase(entit.EntityName) = UCase("Polyline") Then
                Set entPoly = entit
                entPolylines.Add entPoly
            ElseIf UCase(entit.EntityName) = UCase("MText") Then
                Set entMText = entit
                cadMTexts.Add entMText
            ElseIf UCase(entit.EntityName) = UCase("Text") Then
                Set entText = entit
                cadTexts.Add entText
            End If
        End If
    Next entit
    
    
    For Each entPoly In entPolylines
       ' PrintPolyline entPoly
    Next entPoly
    
    For Each entMText In cadMTexts
        'PrintMText entMText
    Next entMText
    
    For Each entText In cadTexts
        'PrintText entText
    Next entText
    For Each entBlock In cadBlocks
        PrintBlockRef entBlock
    Next entBlock
   
End Sub
Public Function PrintPolyline(entPolyline As IntelliCAD.entity)
    Debug.Print "<polyline><layer>" & entPolyline.layer & "</layer>"
    Debug.Print "<color>" & entPolyline.Color & "</color><length>" & Measure(entPolyline) & "</length>"
    Debug.Print PrintVertexes(entPolyline)
    Debug.Print "</polyline>"
End Function
Public Function PrintMText(entMText As IntelliCAD.mText)
    Debug.Print "<mtext><layer>" & entMText.layer & "</layer><textstring>" & entMText.TextString & "</textstring></mtext>"
End Function
Public Function PrintText(entText As IntelliCAD.text)
    Debug.Print "<text><layer>" & entText.layer & "</layer><textstring>" & entText.TextString & "</textstring><insertionpoint><x>" & entText.insertionPoint.x & "</x><y>" & entText.insertionPoint.y & "</y><z>" & entText.insertionPoint.Z & "</z></insertionpoint></text>"
End Function
Public Function PrintBlockRef(entText As IntelliCAD.blockInsert)
    Debug.Print "<blockref><layer>" & entText.layer & "</layer>"
    Debug.Print "   <name>" & entText.name & "</name>"
    Debug.Print "   <insertionpoint><x>" & entText.insertionPoint.x & "</x><y>" & entText.insertionPoint.y & "</y><z>" & entText.insertionPoint.Z & "</z></insertionpoint>"
    Debug.Print "</blockref>"
End Function
Public Function IsOurLayer(layerName As String) As Boolean
    
        If UCase(Mid(layerName, 1, 6)) = UCase("Pipes.") _
            Or UCase(layerName) = UCase("Pipes") _
            Or UCase(layerName) = UCase("Pipe Labels") _
            Or UCase(layerName) = UCase("Tees") _
            Or UCase(layerName) = UCase("TeeLabels") _
            Or UCase(layerName) = UCase("Heads") _
            Or UCase(layerName) = UCase("HeadLabels") _
            Or UCase(layerName) = UCase("Risers") _
            Or UCase(layerName) = UCase("RiserLabels") _
            Or UCase(layerName) = UCase("Floor Connectors") _
            Or UCase(layerName) = UCase("Floor Tags") _
            Or UCase(layerName) = UCase("ElevationBox") Then
            IsOurLayer = True
        Else
            IsOurLayer = False
        End If
End Function
Public Function PrintVertexes(inputPolyline As IntelliCAD.entity)
Dim strPipeRow As String
Dim x As Long
Dim lastX As Long
Dim Coordinates As Variant
    
strPipeRow = strPipeRow & vbTab & vbTab & vbTab & vbTab & "<vertexes>" & vbCrLf
                Coordinates = inputPolyline.Coordinates
                lastX = (Coordinates.Count - 1)
                For x = 0 To lastX Step 2
                        strPipeRow = strPipeRow & vbTab & vbTab & vbTab & vbTab & "<vertex>" & vbCrLf
                        strPipeRow = strPipeRow & vbTab & vbTab & vbTab & vbTab & vbTab & "<x>" & Coordinates(x) & "</x>" & vbCrLf
                        strPipeRow = strPipeRow & vbTab & vbTab & vbTab & vbTab & vbTab & "<y>" & Coordinates(x + 1) & "</y>" & vbCrLf
                        strPipeRow = strPipeRow & vbTab & vbTab & vbTab & vbTab & vbTab & "<z>0</z>" & vbCrLf
                        strPipeRow = strPipeRow & vbTab & vbTab & vbTab & vbTab & "</vertex>" & vbCrLf
                Next x
                strPipeRow = strPipeRow & vbTab & vbTab & vbTab & vbTab & "</vertexes>" & vbCrLf
                PrintVertexes = strPipeRow
End Function
