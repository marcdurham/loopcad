VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Pipe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public startPoint As Variant
Public endPoint As Variant
Public actualLen As Double
Public roundedLen As Double
Public PipeNumber As Long
Public Diameter As Double
Public StartNode As node
Public EndNode As node
Public name As String
Public polyline As IntelliCAD.lwPolyline
Public Vertexes As Collection
Public Property Get RoundedLengthFeet() As Double
    RoundedLengthFeet = roundedLen
End Property
Public Property Let ActualLength(val As Double)
    roundedLen = CDbl(CLng(val / 12)) + 1
    actualLen = val
End Property
Public Property Get ActualLength() As Double
    ActualLength = actualLen
End Property


Public Sub SetPipeVertexes(InputPipe As Pipe)
Dim LastCoordinate As Long
Dim x As Long
Dim v As Vertex
    
    If Not InputPipe.polyline Is Nothing Then
        If InputPipe.Vertexes Is Nothing Then InputPipe.Vertexes = New Collection
        
        LastCoordinate = (InputPipe.polyline.Coordinates.Count - 1)
        For x = 0 To LastCoordinate Step 3
                v.x = InputPipe.polyline.Coordinates(x)
                v.y = InputPipe.polyline.Coordinates(x + 1)
                v.Z = InputPipe.polyline.Coordinates(x + 2)
                InputPipe.Vertexes.Add v
        Next x
        
    End If

End Sub
Public Function GetLength() As Double
    GetLength = Measure(Me.polyline)
End Function
Public Function SetEndPoint(InputPoint As Variant) As Boolean
Dim inputX As Double
Dim inputY As Double
Dim inputZ As Double
Dim endX As Double
Dim endY As Double
Dim endPointCount As Long
Dim newCoords() As Double

    If (InputPoint.Count - 1) = 2 Then
        'Debug.Print "End point for " & Me.Name
        endPointCount = (Me.polyline.Coordinates.Count - 1)
        inputX = InputPoint(0)
        inputY = InputPoint(1)
        endX = Me.polyline.Coordinates(endPointCount - 1)
        endY = Me.polyline.Coordinates(endPointCount)
        inputZ = 0
        
        ReDim newCoords(Me.polyline.Coordinates.Count - 1)
        For x = 0 To (Me.polyline.Coordinates.Count - 1)
            newCoords(x) = Me.polyline.Coordinates(x)
        Next x
        newCoords((Me.polyline.Coordinates.Count - 1) - 1) = inputX
        newCoords((Me.polyline.Coordinates.Count - 1)) = inputY
        Me.polyline.Coordinates = newCoords
        Me.polyline.Update

        
    Else
        Debug.Print "Weird .Count - 1 for endpoint = " & (InputPoint.Count - 1)
    End If
    
End Function
Public Function SetStartPoint(InputPoint As point) As Boolean
Dim inputX As Double
Dim inputY As Double
Dim inputZ As Double
Dim startX As Double
Dim startY As Double
Dim endPointCount As Long
Dim newCoords As IntelliCAD.point


    'If (InputPoint.Count - 1) = 2 Then
        endPointCount = (Me.polyline.Coordinates.Count - 1)
        inputX = InputPoint.x
        inputY = InputPoint.y
        startX = Me.polyline.Coordinates(0).x
        startY = Me.polyline.Coordinates(0).y
        inputZ = 0
        
        ReDim newCoords((Me.polyline.Coordinates.Count - 1))
        For x = 0 To (Me.polyline.Coordinates.Count - 1)
            newCoords(x) = Me.polyline.Coordinates(x)
        Next x
        newCoords.x = inputX
        newCoords.y = inputY
        Me.polyline.Coordinates = newCoords
        Me.polyline.Update
    'Else
    '    Debug.Print "Weird .Count - 1 for Startpoint = " & (InputPoint.Count - 1)
    'End If
    
End Function

