Attribute VB_Name = "JobData"
Public Sub SetJobData()
On Error GoTo ErrTrap
    Dim startPoint As Variant
    Dim Block As IntelliCAD.blockInsert
    Dim strJobNameOld As String
    Dim strLeadNumberOld As String
    Dim strJobNameNew As String
    Dim strLeadNumberNew As String

    Set Block = Job.BlockReference
    
    SetAttrib Block, "CALCULATED_BY_COMPANY", Job.CalculatedByCompany
    SetAttrib Block, "JOB_NUMBER", Job.JobNumber
    SetAttrib Block, "JOB_NAME", Job.JobName
    SetAttrib Block, "JOB_SITE_LOCATION", Job.JobSiteLocation
    SetAttribFromBox Block, "SPRINKER_PIPE_TYPE", Job.SprinklerPipeType
    SetAttribFromBox Block, "SPRINKLER_FITTING_TYPE", Job.SprinklerFittingType
    SetAttribFromBox Block, "SUPPLY_NAME", Job.SupplyName
    SetAttribFromBox Block, "SUPPLY_AVAILABLE_FLOW", Job.SupplyAvailableFlow
    SetAttribFromBox Block, "SUPPLY_STATIC_PRESSURE", Job.SupplyStaticPressure
    SetAttribFromBox Block, "SUPPLY_RESIDUAL_PRESSURE", Job.SupplyResidualPressure
    SetAttribFromBox Block, "SUPPLY_ELEVATION", Job.SupplyElevation
    SetAttribFromBox Block, "SUPPLY_PIPE_LENGTH", Job.SupplyPipeLength
    SetAttribFromBox Block, "SUPPLY_PIPE_TYPE", Job.SupplyPipeType
    SetAttribFromBox Block, "SUPPLY_PIPE_SIZE", Job.SupplyPipeSize
    SetAttribFromBox Block, "SUPPLY_PIPE_INTERNAL_DIAMETER", Job.SupplyPipeInternalDiameter
    SetAttribFromBox Block, "SUPPLY_PIPE_FITTINGS_SUMMARY", Job.SupplyPipeFittingsSummary
    SetAttribFromBox Block, "SUPPLY_PIPE_FITTINGS_EQUIV_LENGTH", Job.SupplyPipeFittingsEquivLength
    SetAttribFromBox Block, "SUPPLY_PIPE_ADD_PRESSURE_LOSS", Job.SupplyPipeAddPressureLoss
    SetAttribFromBox Block, "WATER_FLOW_SWITCH_MAKE_MODEL", Job.WaterFlowSwitchMakeModel
    SetAttribFromBox Block, "WATER_FLOW_SWITCH_PRESSURE_LOSS", Job.WaterFlowSwitchPressureLoss
    SetAttribFromBox Block, "DOMESTIC_FLOW_ADDED", Job.DomesticFlowAdded
    
    CheckAddLayer "Job Data", Colors.vicMagenta, False, True
    Block.layer = "Job Data"
    
ExitProc:
    Exit Sub

ErrTrap:
    If Err.Number = -2147352567 Then 'The user pressed escape
        GoTo ExitProc
    ElseIf Err.Number = -2145320928 Then  'user pressed enter
        GoTo ExitProc
    ElseIf Err.Number = -2145386445 Then 'Can't find block file
        MsgBox "Cannot find the block file", vbCritical, "EditJobData()"
        GoTo ExitProc
    ElseIf Err.Number = -2145386420 Then 'Object erased!
        AddJobData
        GoTo ExitProc
    Else
        MsgBox Err.Number & " " & Err.Description, vbCritical, "EditJobData()"
        GoTo ExitProc
        Resume
    End If
    
End Sub

Public Sub ScanJobData()
Dim Entit As IntelliCAD.entity
Dim blk As IntelliCAD.blockInsert
Dim strJobName As String
Dim strLeadNumber As String
Dim blockCount As Long
Dim strSiteLocation As String
Dim strStaticPressure As String
Dim strResidualPressure As String
Dim strAvailableFlow As String
Dim strMeterElevation As String
Dim strMeterPipeLength As String
Dim strMeterPipeInternalDiameter As String
Dim strCalculatedByCompany As String
Dim i As Integer
    
    For i = 0 To (ActiveDocument.ModelSpace.Count - 1)
        Set Entit = ActiveDocument.ModelSpace(i)
        If Entit.EntityName = "BlockInsert" Then
            Set blk = Entit
            If blk.Name = "JobData" Then
                        
                Set Job = New Job
                Set Job.BlockReference = blk
                Job.CalculatedByCompany = GetAttribOrDefault(blk, "0", "CALCULATED_BY_COMPANY")
                Job.JobName = GetAttribOrDefault(blk, "", "JOB_NAME")
                Job.JobNumber = GetAttribOrDefault(blk, "", "JOB_NUMBER")
                Job.JobSiteLocation = GetAttribOrDefault(blk, "", "JOB_SITE_LOCATION")
                Job.SprinklerPipeType = GetAttribOrDefault(blk, "", "SPRINKLER_PIPE_TYPE")
                Job.SprinklerFittingType = GetAttribOrDefault(blk, "", "SPRINKLER_FITTING_TYPE")
                Job.SupplyName = GetAttribOrDefault(blk, "0", "SUPPLY_NAME")
                Job.SupplyStaticPressure = GetAttribOrDefault(blk, "0", "SUPPLY_STATIC_PRESSURE")
                Job.SupplyResidualPressure = GetAttribOrDefault(blk, "0", "SUPPLY_RESIDUAL_PRESSURE")
                Job.SupplyAvailableFlow = GetAttribOrDefault(blk, "0", "SUPPLY_AVAILABLE_FLOW")
                Job.SupplyElevation = GetAttribOrDefault(blk, "0", "SUPPLY_ELEVATION")
                Job.SupplyPipeLength = GetAttribOrDefault(blk, "0", "SUPPLY_PIPE_LENGTH")
                Job.SupplyPipeSize = GetAttribOrDefault(blk, "0", "SUPPLY_PIPE_SIZE")
                Job.SupplyPipeType = GetAttribOrDefault(blk, "0", "SUPPLY_PIPE_TYPE")
                Job.SupplyPipeInternalDiameter = GetAttribOrDefault(blk, "0", "SUPPLY_PIPE_INTERNAL_DIAMETER")
                Job.SupplyPipeFittingsSummary = GetAttribOrDefault(blk, "0", "SUPPLY_PIPE_FITTINGS_SUMMARY")
                Job.SupplyPipeFittingsEquivLength = GetAttribOrDefault(blk, "0", "SUPPLY_PIPE_FITTINGS_EQUIV_LENGTH")
                Job.SupplyPipeAddPressureLoss = GetAttribOrDefault(blk, "0", "SUPPLY_PIPE_ADD_PRESSURE_LOSS")
                Job.WaterFlowSwitchMakeModel = GetAttribOrDefault(blk, "0", "WATER_FLOW_SWITCH_MAKE_MODEL")
                Job.WaterFlowSwitchPressureLoss = GetAttribOrDefault(blk, "0", "WATER_FLOW_SWITCH_PRESSURE_LOSS")
                Job.DomesticFlowAdded = GetAttribOrDefault(blk, "0", "DOMESTIC_FLOW_ADDED")
    
                blockCount = blockCount + 1
            End If
        End If
    Next i
    
    If blockCount = 0 Then Set Job = Nothing
    
End Sub
Public Function GetAttribOrDefault(blk As IntelliCAD.blockInsert, def As String, ParamArray TagStrings())
    
    'Debug.Print "GetAttribOrDefault: " & UBound(TagStrings)
    'GetAttribOrDefault = EmptyDefault(GetFirstAttrib(blk, TagStrings), Default)
    
Dim val As String
Dim i As Integer
Dim s As String

    Debug.Print "GetAttribOrDefault: " & UBound(TagStrings)
    For i = LBound(TagStrings) To UBound(TagStrings)
        s = TagStrings(i)
        val = GetAttrib(blk, s)
        If Len(val) > 0 Then
            GetAttribOrDefault = val
            Exit Function
        End If
    Next i
    
    GetAttribOrDefault = def
    
End Function
Public Function GetFirstAttrib( _
    blk As IntelliCAD.blockInsert, _
    TagStrings() As Variant) As Variant
    
Dim val As String
Dim i As Integer

    Debug.Print "GetFirstAttrib: " & UBound(TagStrings)
    For i = LBound(TagStrings) To UBound(TagStrings)
        val = GetAttrib(blk, TagStrings(i))
        If Len(v) > 0 Then
            GetFirstAttrib = val
            Exit Function
        End If
    Next i
    
End Function
Public Function EmptyDefault(value As String, Default As Variant) As Variant

    EmptyDefault = IIf(Trim(value) = "", Default, value)
    
End Function

Public Sub EditJobData()
On Error GoTo ErrTrap
    Dim form As frmJobData
    
    Set form = New frmJobData
    form.Show
    'frmJobData.Show
    
    Exit Sub
    
    
    Dim startPoint As Variant
    Dim Block As IntelliCAD.blockInsert
    Dim strJobNameOld As String
    Dim strLeadNumberOld As String
    Dim strJobNameNew As String
    Dim strLeadNumberNew As String
    Dim strSiteLocationOld As String
    Dim strStaticPressureOld As String
    Dim strResidualPressureOld As String
    Dim strAvailableFlowOld As String
    Dim strMeterElevationOld As String
    Dim strMeterPipeLengthOld As String
    Dim strMeterPipeInternalDiameterOld As String
    
    Dim strSiteLocationNew As String
    Dim strStaticPressureNew As String
    Dim strResidualPressureNew As String
    Dim strAvailableFlowNew As String
    Dim strMeterElevationNew As String
    Dim strMeterPipeLengthNew As String
    Dim strMeterPipeInternalDiameterNew As String
    Dim strCalculatedByCompany As String
    
    Set Block = Job.BlockReference
    
    strCalculatedByCompany = GetAttrib(Block, "CALCULATED_BY_COMPANY")
    strJobNameOld = GetAttrib(Block, "JOB_NAME")
    strLeadNumberOld = GetAttrib(Block, "LEAD_NUMBER")
    strSiteLocationOld = GetAttrib(Block, "SITE_LOCATION")
    strStaticPressureOld = GetAttrib(Block, "STATIC_PRESSURE")
    strResidualPressureOld = GetAttrib(Block, "RESIDUAL_PRESSURE")
    strAvailableFlowOld = GetAttrib(Block, "AVAILABLE_FLOW")
    strMeterElevationOld = GetAttrib(Block, "METER_ELEVATION")
    strMeterPipeLengthOld = GetAttrib(Block, "METER_PIPE_LENGTH")
    strMeterPipeInternalDiameterOld = GetAttrib(Block, "METER_PIPE_INTERNAL_DIAMETER")

    
    strJobNameNew = InputBox("Input Job Name", , strJobNameOld)
    If strJobNameNew = "" Then Exit Sub
    SetAttrib Block, "JOB_NAME", strJobNameNew
    
    strLeadNumberNew = InputBox("Input Lead Number", , strLeadNumberOld)
    If strLeadNumberNew = "" Then Exit Sub
    SetAttrib Block, "LEAD_NUMBER", strLeadNumberNew
    
    strSiteLocationNew = InputBox("Site Location", , strSiteLocationOld)
    SetAttrib Block, "SITE_LOCATION", strSiteLocationNew
    
    strStaticPressureNew = InputBox("Static Pressure (psi)", , strStaticPressureOld)
    SetAttribFromBox Block, "STATIC_PRESSURE", strStaticPressureNew
    
    strResidualPressureNew = InputBox("Residual Pressure (psi)", , strResidualPressureOld)
    SetAttribFromBox Block, "RESIDUAL_PRESSURE", strResidualPressureNew
    
    strAvailableFlowNew = InputBox("Available Flow (gpm)", , strAvailableFlowOld)
    SetAttribFromBox Block, "AVAILABLE_FLOW", strAvailableFlowNew
    
    strMeterElevationNew = InputBox("Elevation of meter (ft)", , strMeterElevationOld)
    SetAttribFromBox Block, "METER_ELEVATION", strMeterElevationNew

    
    strMeterPipeLengthNew = InputBox("Length of pipe to meter (ft)", , strMeterPipeLengthOld)
    SetAttribFromBox Block, "METER_PIPE_LENGTH", strMeterPipeLengthNew
    
    strMeterPipeInternalDiameterNew = InputBox("Internal diameter of pipe to meter (inches)", , strMeterPipeInternalDiameterOld)
    SetAttribFromBox Block, "METER_PIPE_INTERNAL_DIAMETER", strMeterPipeInternalDiameterNew
    
    CheckAddLayer "Job Data", Colors.vicMagenta, False, True
    Block.layer = "Job Data"
    
    

ExitProc:
    Exit Sub
    
'        SetAttrib blockHead, "SITE_LOCATION", InputBox("Site Location", , "123 Street, City, State Zip")
'    SetAttrib blockHead, "STATIC_PRESSURE", InputBox("Static Pressure (psi)", , "55")
'    SetAttrib blockHead, "RESIDUAL_PRESSURE", InputBox("Residual Pressure (psi)", , "40")
'    SetAttrib blockHead, "AVAILABLE_FLOW", InputBox("Available Flow (gpm)", , "45")
'    SetAttrib blockHead, "METER_ELEVATION", InputBox("Elevation of Meter (ft)", , "100")
'    SetAttrib blockHead, "METER_PIPE_LENGTH", InputBox("Length of pipe to meter (ft)", , "50")
'    SetAttrib blockHead, "METER_PIPE_INTERNAL_DIAMETER", InputBox("Internal diameter of pipe to meter (inches)", , "0.860")


ErrTrap:
    If Err.Number = -2147352567 Then 'The user pressed escape
        GoTo ExitProc
    ElseIf Err.Number = -2145320928 Then  'user pressed enter
        GoTo ExitProc
    ElseIf Err.Number = -2145386445 Then 'Can't find block file
        MsgBox "Cannot find the block file", vbCritical, "EditJobData()"
        GoTo ExitProc
    ElseIf Err.Number = -2145386420 Then 'Object erased!
        AddJobData
        GoTo ExitProc
    Else
        MsgBox Err.Number & " " & Err.Description, vbCritical, "EditJobData()"
        GoTo ExitProc
        Resume
    End If
    
End Sub

