Attribute VB_Name = "ElevationBoxesTests"
Public Sub TestGetElevation()
    Dim str As String
    Dim p  As Variant
    p = ActiveDocument.Utility.GetPoint(, "Click where you want the test point.")
    str = GetElevationBox(p).elevation
    Debug.Print "You clicked elevation box " & str

End Sub
Public Sub TestGetElevationBoxFloorTag()
    Dim str As String
    Dim p  As Variant
    p = ActiveDocument.Utility.GetPoint(, "Click where you want the test point.")
    str = GetElevationBoxFloorTag(p).name
    Debug.Print "You clicked Floor Tag " & str

End Sub
