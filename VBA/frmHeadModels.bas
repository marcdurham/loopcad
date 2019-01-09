VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmHeadModels 
   Caption         =   "Insert Fire Sprinkler Head"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8220
   OleObjectBlob   =   "frmHeadModels.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmHeadModels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





'Public lHeadCoverage(1 To 3) As String
'Public lHeads(1 To 3) As String

Private Sub buttonCancel_Click()
    Me.Hide
End Sub

Private Sub buttonInsertHead_Click()

    Me.Hide
    'Scanner.InsertHeadModel Me.listHeadModels.Value, Me.listCoverages.Value, Me.listSlopes.Value, Me.listTemperatures.Value
    
    Scanner.InsertHeadModelCode GetCoverage(), GetSide(), GetShortCode() 'GetModelCode()
    
End Sub


Public Sub LoadList()
Dim xHeadModel As HeadModel
Dim headModels As Collection
    
    Set headModels = Scanner.GetHeadModels()
    
    Me.listHeadModels.Clear
    For Each xHeadModel In headModels
        Me.listHeadModels.AddItem xHeadModel.ModelNumber
    Next xHeadModel
    
    If headModels.Count > 0 Then
        Me.listHeadModels.SetFocus
        Me.listHeadModels.value = headModels(1).ModelNumber
    End If
    
    Me.listSides.Clear
    Me.listSides.AddItem "Top"
    Me.listSides.AddItem "Bottom"
    Me.listSides.AddItem "Right"
    Me.listSides.AddItem "Left"
    
End Sub



Private Sub listCoverages_Change()
    ShowModelCode
End Sub



Private Sub listHeadModels_Change()
    LoadHeadModels
    ShowModelCode
End Sub
Public Sub LoadHeadModels()
Dim xHeadModel As HeadModel
Dim xHeadModelItem As HeadModelItem
Dim xTemperature As HeadModelTemperature
Dim xSlope As HeadModelSlope

    Set xHeadModel = Scanner.GetHeadModels(listHeadModels.value)
    
    
    If Not xHeadModel Is Nothing Then
        Me.listTemperatures.Clear
        If Not xHeadModel.Temperatures Is Nothing Then
            For Each xTemperature In xHeadModel.Temperatures
                Me.listTemperatures.AddItem CStr(xTemperature.Temperature)
            Next xTemperature

            If xHeadModel.Temperatures.Count > 0 Then
                If xHeadModel.DefaultTemperature <> 0 Then
                    Me.listTemperatures.SetFocus
                    Me.listTemperatures.value = CStr(xHeadModel.DefaultTemperature)
                Else
                    Me.listTemperatures.SetFocus
                    Me.listTemperatures.value = CStr(xHeadModel.Temperatures(1).Temperature)
                End If
            End If
        End If
        
        Me.listSlopes.Clear
        If Not xHeadModel.Slopes Is Nothing Then
            For Each xSlope In xHeadModel.Slopes
                Me.listSlopes.AddItem xSlope.SlopeName
            Next xSlope
            
            If xHeadModel.Slopes.Count > 0 Then
                Me.listSlopes.SetFocus
                Me.listSlopes.value = xHeadModel.Slopes(1).SlopeName
            End If
        End If
        
        Me.listCoverages.Clear
        If Not xHeadModel.HeadModelItems Is Nothing Then
            If Not xHeadModel.Slopes Is Nothing Then
                If xHeadModel.Slopes.Count > 0 Then
                    For Each xHeadModelItem In xHeadModel.Slopes(1).HeadModelItems
                        Me.listCoverages.AddItem xHeadModelItem.Coverage
                    Next xHeadModelItem
                    
                    If xHeadModel.HeadModelItems.Count > 0 Then
                        Me.listCoverages.SetFocus
                        Me.listCoverages.value = xHeadModel.Slopes(1).HeadModelItems(1).Coverage
                    End If
                End If
            End If
        End If
    End If
    

End Sub
Public Sub LoadHeadModelSlope()
Dim xHeadModel As HeadModel
Dim xHeadModelItem As HeadModelItem
Dim xTemperature As HeadModelTemperature
Dim xSlope As HeadModelSlope

    Set xHeadModel = Scanner.GetHeadModels(listHeadModels.value)
    If xHeadModel.Slopes Is Nothing Then Exit Sub
    If xHeadModel.Slopes.Count = 0 Then Exit Sub
    If IsNull(listSlopes.value) Then Exit Sub
    If listSlopes.value = "" Then Exit Sub
    Set xSlope = Scanner.GetHeadModelSlopes(xHeadModel.ModelNumber)(listSlopes.value)
    
    
    

    Me.listCoverages.Clear
    If Not xHeadModel Is Nothing Then
        If Not xSlope Is Nothing Then
            If Not xSlope.HeadModelItems Is Nothing Then
                For Each xHeadModelItem In xSlope.HeadModelItems
                    Me.listCoverages.AddItem xHeadModelItem.Coverage
                Next xHeadModelItem
                
                If xHeadModel.HeadModelItems.Count > 0 Then Me.listCoverages.value = xSlope.HeadModelItems(1).Coverage
            End If

        End If
    End If
    
    
    
End Sub



Private Sub listHeadModelTemperatures_Change()
Dim xHeadModel As HeadModel
Dim xTemperature As HeadModelTemperature
Dim xHeadModelItem As HeadModelItem


    Set xHeadModel = Scanner.GetHeadModels(listHeadModels.value)
    If Not IsNull(Me.listHeadModelTemperatures.value) Then
        Set xTemperature = xHeadModel.Temperatures(Me.listHeadModelTemperatures.value)
        
        Me.listHeadModelTemperatures.Clear
        If Not xTemperature.Coverages Is Nothing Then
            For Each xHeadModelItem In xTemperature.Coverages
                Me.listHeadModelCoverage.AddItem xHeadModelItem.Coverage
            Next xHeadModelItem
            If xTemperature.Coverages.Count > 0 Then Me.listHeadModelCoverage.value = xTemperature.Coverages(1).Coverage
        End If
    End If

End Sub


Private Sub listSlopes_Change()
    LoadHeadModelSlope
    ShowModelCode
End Sub

Private Sub listTemperatures_Change()
    ShowModelCode
End Sub



Private Sub UserForm_Activate()
    LoadList
End Sub

Public Sub ShowModelCode()
    'Me.labelModelCode.Caption = Me.listHeadModels.Value & "-" & Me.listCoverages.Value & "-" & Me.listSlopes.Value & "-" & Me.listTemperatures.Value
    Me.labelModelCode.Caption = GetModelCode()
    Me.labelShortCode.Caption = GetShortCode()
End Sub
Public Function GetModelCode() As String
Dim model As String
Dim Coverage As String
Dim Slope As String
Dim Temp As String

    model = GetModelNumber
    Coverage = GetCoverage()
    If Coverage = "" Then
        Coverage = ""
    Else
        Coverage = "-" & Coverage
    End If
    If IsNull(Me.listSlopes.value) Then
        Slope = ""
    Else
        Slope = "-" & Me.listSlopes.value
    End If
    If IsNull(Me.listTemperatures.value) Then
        Temp = ""
    Else
        Temp = "-" & Me.listTemperatures.value
    End If
    GetModelCode = model & Coverage & Slope & Temp
End Function
Public Function GetModelNumber() As String
Dim model As String
    If IsNull(Me.listHeadModels.value) Then
        model = ""
    Else
        model = Me.listHeadModels.value
    End If
    GetModelNumber = model
End Function
Public Function GetCoverage() As String
Dim Coverage As String
    If IsNull(Me.listCoverages.value) Then
        Coverage = ""
    Else
        Coverage = Me.listCoverages.value
    End If
    GetCoverage = Coverage
End Function
Public Function GetSide() As String
Dim side As String
    If IsNull(Me.listSides.value) Then
        side = ""
    Else
        side = Me.listSides.value
    End If
    GetSide = side
End Function
Public Function GetShortCode() As String
Dim model As String
Dim Coverage As String
Dim CoverageDepth As String
Dim CoverageWidth As String
Dim Slope As String
Dim Temp As String

    model = GetModelNumber
    Coverage = GetCoverage()
    If Coverage = "" Then
        Coverage = ""
        CoverageDepth = ""
        CoverageWidth = ""
    Else
        CoverageDepth = GetCoverageDepth(Coverage)
        CoverageWidth = GetCoverageWidth(Coverage)
        If CoverageDepth = CoverageWidth Then
            Coverage = "-" & CoverageDepth
        Else
            Coverage = "-" & CoverageDepth & CoverageWidth
        End If
    End If
    If IsNull(Me.listSlopes.value) Then
        Slope = ""
    Else
        Slope = Me.listSlopes.value
        If Slope = "S0" Then
            Slope = ""
        Else
            Slope = Slope
        End If
    End If
    
    If IsNull(Me.listTemperatures.value) Then
        Temp = ""
    Else
        Temp = Me.listTemperatures.value
        If Temp = "155" Then
            Temp = ""
        Else
            Temp = "-" & Temp
        End If
    End If
    
    GetShortCode = model & Coverage & Slope & Temp
End Function

