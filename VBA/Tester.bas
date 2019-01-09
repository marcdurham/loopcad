Attribute VB_Name = "Tester"
Public Sub Test()
On Error GoTo ErrTrap

    Dim c As New LoopCADPTest1.Class1
    
    Set c.Doc = ActiveDocument
    
    c.Main

ExitSub:
    Exit Sub

ErrTrap:
    MsgBox "LoopCAD Error: " & Err.Description, vbOKOnly, "LoopCAD Error"
    Resume ExitSub
    
End Sub
Public Sub ChangeHeadModel(FromModel As String, ToModel As String)
    Dim entity As IntelliCAD.entity
    Dim block As IntelliCAD.blockInsert
    Dim newBlock As IntelliCAD.blockInsert
    Dim i As Long
    
    i = 0
    
    For Each entity In ActiveDocument.ModelSpace
        If entity.EntityName = "BlockInsert" Then
            Set block = entity
            If block.name = FromModel Then
                Set newBlock = ActiveDocument.ModelSpace.InsertBlock(block.insertionPoint, ToModel, 1, 1, 1, 0)
                newBlock.layer = "0a__Heads"
                block.Delete
                
                Debug.Print "Head Replaced " & i & " " & FromModel & " to " & ToModel
                i = i + 1
            End If
        End If
    Next entity
    
     Debug.Print "Head Replacment done"
    
End Sub
Public Sub listHeadModels(model As String)
    Dim entity As IntelliCAD.entity
    Dim block As IntelliCAD.blockInsert
    Dim newBlock As IntelliCAD.blockInsert
    Dim i As Long
    
    i = 0
    
    For Each entity In ActiveDocument.ModelSpace
        If entity.EntityName = "BlockInsert" Then
            Set block = entity
            If UCase(block.name) Like UCase(model) Then
                Debug.Print "Head " & i & " " & block.name
                i = i + 1
            End If
        End If
    Next entity
    
     Debug.Print "** End of head list **"
    
End Sub
Public Sub GetMidPointTest()
Dim Result(0 To 2) As Double
Dim vResult As Variant
Dim failCount As Long
    failCount = 0

    vResult = GetMidPoint(0, -10, 10, 0)
    If vResult(0) = 5 And vResult(1) = -5 Then
        Debug.Print "OK"
    Else
        Debug.Print "FAIL " & vResult(0) & ", " & vResult(1)
        failCount = failCount + 1
    End If
    
    vResult = GetMidPoint(10, 0, 0, -10) ' reverse from first one
    If vResult(0) = 5 And vResult(1) = -5 Then
        Debug.Print "OK"
    Else
        Debug.Print "FAIL " & vResult(0) & ", " & vResult(1)
        failCount = failCount + 1
    End If
    
    vResult = GetMidPoint(0, 10, 10, 0)
    If vResult(0) = 5 And vResult(1) = 5 Then
        Debug.Print "OK"
    Else
       Debug.Print "FAIL " & vResult(0) & ", " & vResult(1)
       failCount = failCount + 1
    End If

    vResult = GetMidPoint(-10, 0, 0, -10)
    If vResult(0) = -5 And vResult(1) = -5 Then
        Debug.Print "OK"
    Else
        Debug.Print "FAIL " & vResult(0) & ", " & vResult(1)
        failCount = failCount + 1
    End If

    vResult = GetMidPoint(-10, 0, 0, 10)
    If vResult(0) = -5 And vResult(1) = 5 Then
        Debug.Print "OK"
    Else
        Debug.Print "FAIL " & vResult(0) & ", " & vResult(1)
        failCount = failCount + 1
    End If
    
    vResult = GetMidPoint(-5, 5, 5, -5)
    If vResult(0) = 0 And vResult(1) = 0 Then
        Debug.Print "OK"
    Else
        Debug.Print "FAIL " & vResult(0) & ", " & vResult(1)
        failCount = failCount + 1
    End If
    
    vResult = GetMidPoint(-5, -5, 5, 5)
    If vResult(0) = 0 And vResult(1) = 0 Then
        Debug.Print "OK"
    Else
        Debug.Print "FAIL " & vResult(0) & ", " & vResult(1)
        failCount = failCount + 1
    End If
    
    vResult = GetMidPoint(0, -5, 0, 5)
    If vResult(0) = 0 And vResult(1) = 0 Then
        Debug.Print "OK"
    Else
        Debug.Print "FAIL " & vResult(0) & ", " & vResult(1)
        failCount = failCount + 1
    End If
    
    vResult = GetMidPoint(-5, 0, 5, 0)
    If vResult(0) = 0 And vResult(1) = 0 Then
        Debug.Print "OK"
    Else
        Debug.Print "FAIL " & vResult(0) & ", " & vResult(1)
        failCount = failCount + 1
    End If
    
    vResult = GetMidPoint(-5, 5, 5, 5)
    If vResult(0) = 0 And vResult(1) = 5 Then
        Debug.Print "OK"
    Else
        Debug.Print "FAIL " & vResult(0) & ", " & vResult(1)
        failCount = failCount + 1
    End If
    
    vResult = GetMidPoint(-5, -5, 5, -5)
    If vResult(0) = 0 And vResult(1) = -5 Then
        Debug.Print "OK"
    Else
        Debug.Print "FAIL " & vResult(0) & ", " & vResult(1)
        failCount = failCount + 1
    End If
    
    vResult = GetMidPoint(0, 0, 10, 10)
    If vResult(0) = 5 And vResult(1) = 5 Then
        Debug.Print "OK"
    Else
        Debug.Print "FAIL " & vResult(0) & ", " & vResult(1)
        failCount = failCount + 1
    End If
    
    vResult = GetMidPoint(0, 0, -10, -10)
    If vResult(0) = -5 And vResult(1) = -5 Then
        Debug.Print "OK"
    Else
        Debug.Print "FAIL " & vResult(0) & ", " & vResult(1)
        failCount = failCount + 1
    End If
    
    vResult = GetMidPoint(-10, -10, 0, 0)
    If vResult(0) = -5 And vResult(1) = -5 Then
        Debug.Print "OK"
    Else
        Debug.Print "FAIL " & vResult(0) & ", " & vResult(1)
        failCount = failCount + 1
    End If
    
    If failCount > 0 Then
        Debug.Print "Tests Failed: " & failCount
    Else
        Debug.Print "Tests All Passed"
    End If
    
End Sub

Public Sub OpenFSDXTest()
    
    Dim con As ADODB.Connection
    Dim rs As ADODB.Recordset
    
    rs.Open "dude.xml"
    
    
    
End Sub
Public Function RecordsetFromXMLString(sXML As String) As Recordset

    Dim oStream As ADODB.Stream
    Set oStream = New ADODB.Stream
    
    oStream.Open
    oStream.WriteText sXML   'Give the XML string to the ADO Stream
     'oStream.Open "file://C:/ActiveDocuments%20and%20Settings/Marc/My%20ActiveDocuments/Code/Hydronics/Hydronics.Dev/Hydronics/bin/x86/Debug/Spec%20Files/tee_specs.xml", adModeRead, adOpenStreamAsync
    

    oStream.Position = 0    'Set the stream position to the start

    Dim oRecordset As ADODB.Recordset
    Set oRecordset = New ADODB.Recordset
    

    oRecordset.Open "C:\ActiveDocuments and Settings\Marc\My ActiveDocuments\Code\Hydronics\Hydronics.Dev\Hydronics\bin\x86\Debug\Spec Files\tee_specs.xml"
    oRecordset.Open "file://C:/ActiveDocuments%20and%20Settings/Marc/My%20ActiveDocuments/Code/Hydronics/Hydronics.Dev/Hydronics/bin/x86/Debug/Spec%20Files/tee_specs.xml"
    'oRecordset.Open oStream    'Open a recordset from the stream

    oStream.Close
    Set oStream = Nothing

    Set RecordsetFromXMLString = oRecordset  'Return the recordset

    Set oRecordset = Nothing

End Function
Public Sub TestInLineBox()
    If InLineBox(-5, 5, 5, -5, 0, 0) Then
        Debug.Print "PASS TestInLineBox Test 1"
    Else
        Debug.Print "FAIL TestInLineBox Test 1"
    End If
    
    If InLineBox(-5, 5, 5, -5, -3, 0) Then
        Debug.Print "PASS TestInLineBox Test 2"
    Else
        Debug.Print "FAIL TestInLineBox Test 2"
    End If
    
    If InLineBox(5, -5, -5, 5, 3, 3) Then
        Debug.Print "PASS TestInLineBox Test 3"
    Else
        Debug.Print "FAIL TestInLineBox Test 3"
    End If
    
    If InLineBox(0, 0, -5, -5, 0, 0) Then
        Debug.Print "PASS TestInLineBox Test 4"
    Else
        Debug.Print "FAIL TestInLineBox Test 4"
    End If
    
    If InLineBox(-5, -5, 0, 0, 0, 0) Then
        Debug.Print "PASS TestInLineBox Test 5"
    Else
        Debug.Print "FAIL TestInLineBox Test 5"
    End If
    
    If Not InLineBox(-5, 5, 5, -5, -6, -6) Then
        Debug.Print "PASS TestInLineBox Test 6"
    Else
        Debug.Print "FAIL TestInLineBox Test 6"
    End If
    
    If Not InLineBox(-5, 5, 5, -5, -6, 0) Then
        Debug.Print "PASS TestInLineBox Test 7"
    Else
        Debug.Print "FAIL TestInLineBox Test 7"
    End If
    
    
End Sub

Public Sub TestGetLineRatio()
    If GetLineY(-10, 10, 7.5, 1.25, -5) = 7.5 Then
        Debug.Print "PASS TestGetLineRatio Test 1 = 7.5"
    Else
        Debug.Print "FAIL TestGetLineRatio Test 1 = 7.5"
    End If
    
    If GetLineY(-10, 10, 7.5, 1.25, -10) = 10 Then
        Debug.Print "PASS TestGetLineRatio Test 2 = 10"
    Else
        Debug.Print "FAIL TestGetLineRatio Test 2 = 10"
    End If
    
    If GetLineY(-10, 10, 7.5, 1.25, 0) = 5 Then
        Debug.Print "PASS TestGetLineRatio Test 3 = 5"
    Else
        Debug.Print "FAIL TestGetLineRatio Test 3 = 5"
    End If

    If GetLineY(-10, 10, 7.5, 1.25, 5) = 2.5 Then
        Debug.Print "PASS TestGetLineRatio Test 4 = 2.5"
    Else
        Debug.Print "FAIL TestGetLineRatio Test 4 = 2.5"
    End If
    
    If GetLineY(-10, 10, 7.5, 1.25, 7.5) = 1.25 Then
        Debug.Print "PASS TestGetLineRatio Test 5 = 1.25"
    Else
        Debug.Print "FAIL TestGetLineRatio Test 5 = 1.25"
    End If
    
    If GetLineY(-10, 10, 7.5, 1.25, 10) = 0 Then
        Debug.Print "PASS TestGetLineRatio Test 6 = 0"
    Else
        Debug.Print "FAIL TestGetLineRatio Test 6 = 0"
    End If
    
    If GetLineY(7.5, 1.25, -10, 10, 10) = 0 Then
        Debug.Print "PASS TestGetLineRatio Test 7 (backwards 6) = 0"
    Else
        Debug.Print "FAIL TestGetLineRatio Test 7 (backwards 6) = 0"
    End If
    
    
    If GetLineY(7.5, 1.25, -10, 10, 7.5) = 1.25 Then
        Debug.Print "PASS TestGetLineRatio Test 8 (backwards 5) = 1.25"
    Else
        Debug.Print "FAIL TestGetLineRatio Test 8 (backwards 5) = 1.25"
    End If
    
    If GetLineY(7.5, 1.25, -10, 10, 5) = 2.5 Then
        Debug.Print "PASS TestGetLineRatio Test 9 (backwards 4) = 2.5"
    Else
        Debug.Print "FAIL TestGetLineRatio Test 9 (backwards 4) = 2.5"
    End If
    
    If GetLineY(7.5, 1.25, -10, 10, 0) = 5 Then
        Debug.Print "PASS TestGetLineRatio Test 10 (backwards 3)  = 5"
    Else
        Debug.Print "FAIL TestGetLineRatio Test 10 (backwards 3)  = 5"
    End If
    
    If GetLineY(7.5, 1.25, -10, 10, -10) = 10 Then
        Debug.Print "PASS TestGetLineRatio Test 11 (backwards 2) = 10 "
    Else
        Debug.Print "FAIL TestGetLineRatio Test 11 (backwards 2) = 10 "
    End If
    
    If GetLineY(7.5, 1.25, -10, 10, -5) = 7.5 Then
        Debug.Print "PASS TestGetLineRatio Test 12 (backwards 1) = 7.5"
    Else
        Debug.Print "FAIL TestGetLineRatio Test 12 (backwards 1) = 7.5"
    End If
    

    
    If GetLineY(5, 5, -5, -5, 0) = 0 Then
        Debug.Print "PASS TestGetLineRatio Test 2.0 (new tests) = 0"
    Else
        Debug.Print "FAIL TestGetLineRatio Test 2.0 (new tests) = 0"
    End If
    
    If GetLineY(-5, -5, 5, 5, 0) = 0 Then
        Debug.Print "PASS TestGetLineRatio Test 2.1 (new tests) = 0"
    Else
        Debug.Print "FAIL TestGetLineRatio Test 2.1 (new tests) = 0"
    End If
    
    If GetLineY(-5, 5, 5, -5, 0) = 0 Then
        Debug.Print "PASS TestGetLineRatio Test 3.0 = 0 (new tests) "
    Else
        Debug.Print "FAIL TestGetLineRatio Test 3.0 = 0 (new tests) "
    End If
    
    If GetLineY(5, -5, -5, 5, 0) = 0 Then
        Debug.Print "PASS TestGetLineRatio Test 3.1 = 0 (new tests) "
    Else
        Debug.Print "FAIL TestGetLineRatio Test 3.1 = 0 (new tests) "
    End If
    
    If GetLineY(-10, 10, 7.5, 1.25, 10) = 0 Then
        Debug.Print "PASS TestGetLineRatio Test 4.0 = 5"
    Else
        Debug.Print "FAIL TestGetLineRatio Test 4.0 = 5"
    End If
    
    If GetLineY(-15, -15, -5, -5, -10) = -10 Then
        Debug.Print "PASS TestGetLineRatio Test 5.0 = -10 (All negative)"
    Else
        Debug.Print "FAIL TestGetLineRatio Test 5.0 = -10 (All negative)"
    End If
    
    If GetLineY(0, 0, 0, 0, 0) = 0 Then
        Debug.Print "PASS TestGetLineRatio Test 5.1.0 = 0 (all zero, same point)"
    Else
        Debug.Print "FAIL TestGetLineRatio Test 5.1.0 = 0 (all zero, same point)"
    End If
    
    If GetLineY(0, 0, 0, 0, 1) = "Error: Different point!" Then
        Debug.Print "PASS TestGetLineRatio Test 5.1.1 = 0 (all zero, error different point)"
    Else
        Debug.Print "FAIL TestGetLineRatio Test 5.1.1 = 0 (all zero, error different point)"
    End If
    
    If GetLineY(0, 0, 10, 0, 5) = 0 Then
        Debug.Print "PASS TestGetLineRatio Test 5.2.0 = 0 (flat horizontal)"
    Else
        Debug.Print "FAIL TestGetLineRatio Test 5.2.0 = 0 (flat horizontal)"
    End If
    
    If GetLineY(0, 0, 10, 0, 20) = 0 Then
        Debug.Print "PASS TestGetLineRatio Test 5.2.1 = 0 (flat horizontal, beyond end points)"
    Else
        Debug.Print "FAIL TestGetLineRatio Test 5.2.1 = 0 (flat horizontal, beyond end points)"
    End If
    
    If GetLineY(0, 0, 10, 0, -20) = 0 Then
        Debug.Print "PASS TestGetLineRatio Test 5.2.2 = 0 (flat horizontal, negative X)"
    Else
        Debug.Print "FAIL TestGetLineRatio Test 5.2.2 = 0 (flat horizontal, negative X)"
    End If
    
    If GetLineY(-10, 0, 10, 0, 5) = 0 Then
        Debug.Print "PASS TestGetLineRatio Test 5.3.0 = 0 (flat horizontal, longer)"
    Else
        Debug.Print "FAIL TestGetLineRatio Test 5.3.0 = 0 (flat horizontal, longer)"
    End If
    
    If GetLineY(0, 0, 10, 0, 5) = 0 Then
        Debug.Print "PASS TestGetLineRatio Test 5.3.1 = 0 (flat horizontal, from zero)"
    Else
        Debug.Print "FAIL TestGetLineRatio Test 5.3.1 = 0 (flat horizontal, from zero)"
    End If
    
    If GetLineY(0, -10, 0, 10, 0) = "Infinity" Then
        Debug.Print "PASS TestGetLineRatio Test 5.4 = Infinity (flat vertical, longer)"
    Else
        Debug.Print "FAIL TestGetLineRatio Test 5.4 = Infinity (flat vertical, longer)"
    End If
    
    If GetLineY(0, 0, 0, 10, 0) = "Infinity" Then
        Debug.Print "PASS TestGetLineRatio Test 5.5 = Infinity (flat vertical, from zero) "
    Else
        Debug.Print "FAIL TestGetLineRatio Test 5.5 = Infinity (flat vertical, from zero) "
    End If
    
    If GetLineY(0, -10, 0, 10, 5) = "Error: Point not on line!" Then
        Debug.Print "PASS TestGetLineRatio Test 5.6 = Fail (flat vertical, fail) "
    Else
        Debug.Print "FAIL TestGetLineRatio Test 5.6 = Fail (flat vertical, fail) "
    End If
    
End Sub
Public Sub TestLinePointXY()
    If Not LinePointXY(-10, 10, 7.5, 1.25, 10, 0) Then
        Debug.Print "PASS TestLinePointXY 1"
    Else
        Debug.Print "FAIL TestLinePointXY 1"
    End If
    
    If LinePointXY(-10, 10, 10, 0, 7.5, 1.25) Then
        Debug.Print "PASS TestLinePointXY 2"
    Else
        Debug.Print "FAIL TestLinePointXY 2"
    End If
    
    If LinePointXY(-15, -15, -5, -5, -10, -10) Then
        Debug.Print "PASS TestLinePointXY 3"
    Else
        Debug.Print "FAIL TestLinePointXY 3"
    End If
    
    If LinePointXY(0, 0, 0, 0, 0, 0) Then
        Debug.Print "PASS TestLinePointXY 4.0 (point, on it)"
    Else
        Debug.Print "FAIL TestLinePointXY 4.0 (point, on it)"
    End If
    
    If LinePointXY(-5, 5, -5, 5, -5, 5) Then
        Debug.Print "PASS TestLinePointXY 4.1 (point, on it)"
    Else
        Debug.Print "FAIL TestLinePointXY 4.1 (point, on it)"
    End If
    
    If Not LinePointXY(0, 0, 0, 0, 1, 0) Then
        Debug.Print "PASS TestLinePointXY 5.0 (point, not on it)"
    Else
        Debug.Print "FAIL TestLinePointXY 5.0 (point, not on it)"
    End If
    
    If Not LinePointXY(-5, 5, -5, 5, 1, 0) Then
        Debug.Print "PASS TestLinePointXY 5.1 (point, not on it)"
    Else
        Debug.Print "FAIL TestLinePointXY 5.1 (point, not on it)"
    End If
    
    If Not LinePointXY(0, 0, 0, 0, 1, -2) Then
        Debug.Print "PASS TestLinePointXY 6 (point, not on it)"
    Else
        Debug.Print "FAIL TestLinePointXY 6 (point, not on it)"
    End If
    
    If LinePointXY(-10, 0, 10, 0, 5, 0) Then
        Debug.Print "PASS TestLinePointXY 7 (horizontal line)"
    Else
        Debug.Print "FAIL TestLinePointXY 7(horizontal line)"
    End If

    If Not LinePointXY(-10, 0, 10, 0, 5, -5) Then
        Debug.Print "PASS TestLinePointXY 8 (horizontal line)"
    Else
        Debug.Print "FAIL TestLinePointXY 8(horizontal line, not on it)"
    End If
    
    If LinePointXY(0, -10, 0, 10, 0, 5) Then
        Debug.Print "PASS TestLinePointXY 9 (vertical line, on it)"
    Else
        Debug.Print "FAIL TestLinePointXY 9 (vertical line, on it)"
    End If
    
    If Not LinePointXY(0, -10, 0, 10, 0, 15) Then
        Debug.Print "PASS TestLinePointXY 10 (vertical line, not on it)"
    Else
        Debug.Print "FAIL TestLinePointXY 10 (vertical line, not on it)"
    End If
End Sub

