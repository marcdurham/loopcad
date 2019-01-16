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
    'SetAttrib block, "LEAD_NUMBER", Job.LeadN
    'SetAttrib block, "SITE_LOCATION", Job.SiteLocation
    SetAttrib Block, "JOB_SITE_LOCATION", Job.JobSiteLocation
    
    SetAttribFromBox Block, "SPRINKER_PIPE_TYPE", Job.SprinklerPipeType
    SetAttribFromBox Block, "SPRINKLER_FITTING_TYPE", Job.SprinklerFittingType
    SetAttribFromBox Block, "SUPPLY_NAME", Job.SupplyName
    SetAttribFromBox Block, "SUPPLY_AVAILABLE_FLOW", Job.SupplyAvailableFlow
    SetAttribFromBox Block, "SUPPLY_STATIC_PRESSURE", Job.SupplyStaticPressure
    SetAttribFromBox Block, "SUPPLY_RESIDUAL_PRESSURE", Job.SupplyResidualPressure
    SetAttribFromBox Block, "SUPPLY_ELEVATION", Job.SupplyElevation
    SetAttribFromBox Block, "SUPPLY_MANIFOLD_PIPE_LENGTH", Job.SupplyManifoldPipeLength
    SetAttribFromBox Block, "SUPPLY_MANIFOLD_PIPE_TYPE", Job.SupplyManifoldPipeType
    SetAttribFromBox Block, "SUPPLY_MANIFOLD_PIPE_SIZE", Job.SupplyManifoldPipeSize
    SetAttribFromBox Block, "SUPPLY_MANIFOLD_PIPE_INTERNAL_DIAMETER", Job.SupplyManifoldPipeInternalDiameter
    SetAttribFromBox Block, "SUPPLY_MANIFOLD_FITTINGS_SUMMARY", Job.SupplyManifoldFittingsSummary
    SetAttribFromBox Block, "SUPPLY_MANIFOLD_FITTINGS_EQUIV_LENGTH", Job.SupplyManifoldFittingsEquivLength
    SetAttribFromBox Block, "SUPPLY_MANIFOLD_WATER_FLOW_PRESSURE_LOSS", Job.SupplyManifoldWaterFlowPressureLoss
    SetAttribFromBox Block, "WATER_FLOW_MAKE_MODEL", Job.WaterFlowMakeModel
    SetAttribFromBox Block, "DOMESTIC_FLOW_ADDED", Job.DomesticFlowAdded
    
    'SetAttribFromBox block, "STATIC_PRESSURE", Job.StaticPressure
    'SetAttribFromBox block, "RESIDUAL_PRESSURE", Job.ResidualPressure
    'SetAttribFromBox block, "AVAILABLE_FLOW", Job.AvailableFlow
    'SetAttribFromBox block, "METER_ELEVATION", Job.MeterElevation
    'SetAttribFromBox block, "METER_PIPE_LENGTH", Job.MeterPipeLength
    'SetAttribFromBox block, "METER_PIPE_INTERNAL_DIAMETER", Job.MeterPipeInternalDiameter
    
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
Dim entit As IntelliCAD.entity
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
        Set entit = ActiveDocument.ModelSpace(i)
        If entit.EntityName = "BlockInsert" Then
            Set blk = entit
            If blk.Name = "JobData" Then
              
              
'    SetAttrib block, "CALCULATED_BY_COMPANY", Job.CalculatedByCompany
'    SetAttrib block, "JOB_NUMBER", Job.JobNumber
'    SetAttrib block, "JOB_NAME", Job.JobName
'    'SetAttrib block, "LEAD_NUMBER", Job.LeadN
'    'SetAttrib block, "SITE_LOCATION", Job.SiteLocation
'    SetAttrib block, "JOB_SITE_LOCATION", Job.JobSiteLocation
'
'    SetAttribFromBox block, "SPRINKER_PIPE_TYPE", Job.SprinklerPipeType
'    SetAttribFromBox block, "SPRINKLER_FITTING_TYPE", Job.SprinklerFittingType
'    SetAttribFromBox block, "SUPPLY_NAME", Job.SupplyName
'    SetAttribFromBox block, "SUPPLY_AVAILABLE_FLOW", Job.SupplyAvailableFlow
'    SetAttribFromBox block, "SUPPLY_STATIC_PRESSURE", Job.SupplyStaticPressure
'    SetAttribFromBox block, "SUPPLY_RESIDUAL_PRESSURE", Job.SupplyResidualPressure
'    SetAttribFromBox block, "SUPPLY_ELEVATION", Job.SupplyElevation
'    SetAttribFromBox block, "SUPPLY_MANIFOLD_PIPE_LENGTH", Job.SupplyManifoldPipeLength
'    SetAttribFromBox block, "SUPPLY_MANIFOLD_PIPE_TYPE", Job.SupplyManifoldPipeType
'    SetAttribFromBox block, "SUPPLY_MANIFOLD_PIPE_SIZE", Job.SupplyManifoldPipeSize
'    SetAttribFromBox block, "SUPPLY_MANIFOLD_PIPE_INTERNAL_DIAMETER", Job.SupplyManifoldPipeInternalDiameter
'    SetAttribFromBox block, "SUPPLY_MANIFOLD_FITTINGS_SUMMARY", Job.SupplyManifoldFittingsSummary
'    SetAttribFromBox block, "SUPPLY_MANIFOLD_FITTINGS_EQUIV_LENGTH", Job.SupplyManifoldFittingsEquivLength
'    SetAttribFromBox block, "SUPPLY_MANIFOLD_WATER_FLOW_PRESSURE_LOSS", Job.SupplyManifoldWaterFlowPressureLoss
'    SetAttribFromBox block, "WATER_FLOW_MAKE_MODEL", Job.WaterFlowMakeModel
'    SetAttribFromBox block, "DOMESTIC_FLOW_ADDED", Job.DomesticFlowAdded
'
'    'SetAttribFromBox block, "STATIC_PRESSURE", Job.StaticPressure
'    'SetAttribFromBox block, "RESIDUAL_PRESSURE", Job.ResidualPressure
'    'SetAttribFromBox block, "AVAILABLE_FLOW", Job.AvailableFlow
'    'SetAttribFromBox block, "METER_ELEVATION", Job.MeterElevation
'    'SetAttribFromBox block, "METER_PIPE_LENGTH", Job.MeterPipeLength
'    'SetAttribFromBox block, "METER_PIPE_INTERNAL_DIAMETER", Job.MeterPipeInternalDiameter
'
    
                strJobName = GetAttrib(blk, "JOB_NAME")
                strLeadNumber = GetAttrib(blk, "JOB_NUMBER") 'LEAD_NUMBER
                
                strSiteLocation = GetFirstAttrib(blk, "JOB_SITE_LOCATION")
                strStaticPressure = GetFirstAttrib(blk, "SUPPLY_STATIC_PRESSURE")
                strResidualPressure = GetFirstAttrib(blk, "SUPPLY_RESIDUAL_PRESSURE")
                strAvailableFlow = GetFirstAttrib(blk, "SUPPLY_AVAILABLE_FLOW")
                strMeterElevation = GetFirstAttrib(blk, "SUPPLY_ELEVATION")
                strMeterPipeLength = GetFirstAttrib(blk, "SUPPLY_MANIFOLD_PIPE_LENGTH")
                strMeterPipeInternalDiameter = GetFirstAttrib(blk, "SUPPLY_MANIFOLD_PIPE_INTERNAL_DIAMETER")
                strCalculatedByCompany = GetFirstAttrib(blk, "CALCULATED_BY_COMPANY")
                
                Set Job = New Job
                Set Job.BlockReference = blk
                Job.JobName = strJobName
                Job.JobNumber = strLeadNumber
                Job.JobSiteLocation = strSiteLocation
                Job.SupplyStaticPressure = EmptyDefault(strStaticPressure, 0)
                Job.SupplyResidualPressure = EmptyDefault(strResidualPressure, 0)
                Job.SupplyAvailableFlow = EmptyDefault(strAvailableFlow, 0)
                Job.SupplyElevation = EmptyDefault(strMeterElevation, 0)
                Job.SupplyManifoldPipeLength = EmptyDefault(strMeterPipeLength, 0)
                Job.SupplyManifoldPipeInternalDiameter = EmptyDefault(strMeterPipeInternalDiameter, 0)
                Job.CalculatedByCompany = EmptyDefault(strCalculatedByCompany, 0)

               
                blockCount = blockCount + 1
            End If
        End If
    Next i
    
    If blockCount = 0 Then Set Job = Nothing
    
End Sub
Public Function GetFirstAttrib(blk As IntelliCAD.blockInsert, ParamArray tags())
Dim v As String
Dim tag As Variant

    For Each tag In tags
        v = GetAttrib(blk, tag)
        If Len(v) > 0 Then
            GetFirstAttrib = v
            Exit Function
        End If
    Next tag
    
End Function
Public Function EmptyDefault(value As String, default As String)
    EmptyDefault = IIf(Trim(value) = "", default, value)
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
    
    strJobNameOld = GetAttrib(Block, "JOB_NAME")
    strLeadNumberOld = GetAttrib(Block, "LEAD_NUMBER")
    strSiteLocationOld = GetAttrib(Block, "SITE_LOCATION")
    strStaticPressureOld = GetAttrib(Block, "STATIC_PRESSURE")
    strResidualPressureOld = GetAttrib(Block, "RESIDUAL_PRESSURE")
    strAvailableFlowOld = GetAttrib(Block, "AVAILABLE_FLOW")
    strMeterElevationOld = GetAttrib(Block, "METER_ELEVATION")
    strMeterPipeLengthOld = GetAttrib(Block, "METER_PIPE_LENGTH")
    strMeterPipeInternalDiameterOld = GetAttrib(Block, "METER_PIPE_INTERNAL_DIAMETER")
    strCalculatedByCompany = GetAttrib(Block, "CALCULATED_BY_COMPANY")
    
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

