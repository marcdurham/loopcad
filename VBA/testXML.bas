Attribute VB_Name = "testXML"

Public Function SaveFSDXML()
' This procedure creates XML ActiveDocument
' and saves it to disk.
' Requires msxml.dll (Go to Project --> References and
' and choose Microsoft XML version 2.0, or whatever the
' current version you have installed)
' The example given below will write the following XML
' ActiveDocuments.
'
' <Family>
'    <Member Relationship="Father">
'       <Name>Some Guy</Name>
'    </Member>
' </Family>
'
'but it should be clear how to modify the code
'to create your own ActiveDocuments


   
   Dim objDom As DOMActiveDocument
   Dim objRootElem As IXMLDOMElement
   Dim objMemberElem As IXMLDOMElement
   Dim objMemberRel As IXMLDOMAttribute
   Dim objMemberName As IXMLDOMElement
   
   Set objDom = New DOMActiveDocument
   
   ' Creates root element
   Set rootFile = objDom.createElement("fsdx_file")
   objDom.appendChild rootFile
   
   ' Creates Member element
   Set elementCreation = objDom.createElement("creation")
   rootFile.appendChild elementCreation
   
   Set elementCreated = objDom.createElement("created")
   elementCreation.appendChild elementCreation
   elementCreated.text = Format(Now(), "YYYY-MM-DD") & "T" & Format(Now(), "HH:mm:ss")
   
   
   
   
   
   
   
   
   
   
   Set elementFSSystem = objDom.createElement("fire_sprinkler_system")
   rootFile.appendChild elementFSSystem

   Set elementJob = objDom.createElement("job")
   rootFile.appendChild elementJob

   Set elementNode = objDom.createElement("node")
   rootFile.appendChild elementNode
   
   Set elementPipe = objDom.createElement("pipe")
   rootFile.appendChild elementPipe
   
   Set elementHeadGroup = objDom.createElement("head_group")
   rootFile.appendChild elementHeadGroup
   
   
   ' Create element under Member element, and
   ' gives value "some guy"
   Set objMemberName = objDom.createElement("Name")
   objMemberElem.appendChild objMemberName
   objMemberName.text = "Some ""Guy"" & <hello> 'world'"

   ' Saves XML data to disk.
   objDom.Save ("\\TSCLIENT\C\temp\andrew.xml")
    
End Function

