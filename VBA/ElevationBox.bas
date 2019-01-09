VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ElevationBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Box As IntelliCAD.lwPolyline
Public Box2d As IntelliCAD.polyline
Public elevation As Double
Public InnerBoxes As Collection

Public Static Sub Draw()
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
    
    Set startPoint = ActiveDocument.Utility.GetPoint(, "\n\nClick the first corner for the Elevation Box")
    Set endPoint = ActiveDocument.Utility.GetCorner(startPoint, "\nClick the other corner for the Elevation Box")
    
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


