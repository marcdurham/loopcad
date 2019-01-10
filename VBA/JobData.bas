Attribute VB_Name = "JobData"
Public Sub SetJobData()
On Error GoTo ErrTrap
    Dim startPoint As Variant
    Dim block As IntelliCAD.blockInsert
    Dim strJobNameOld As String
    Dim strLeadNumberOld As String
    Dim strJobNameNew As String
    Dim strLeadNumberNew As String

    
    Set block = Job.BlockReference
    
    SetAttrib block, "CALCULATED_BY_COMPANY", Job.CalculatedByCompany
    SetAttrib block, "JOB_NUMBER", Job.JobNumber
    SetAttrib block, "JOB_NAME", Job.JobName
    'SetAttrib block, "LEAD_NUMBER", Job.LeadN
    'SetAttrib block, "SITE_LOCATION", Job.SiteLocation
    SetAttrib block, "JOB_SITE_LOCATION", Job.JobSiteLocation
    
    SetAttribFromBox block, "SPRINKER_PIPE_TYPE", Job.SprinklerPipeType
    SetAttribFromBox block, "SPRINKLER_FITTING_TYPE", Job.SprinklerFittingType
    SetAttribFromBox block, "SUPPLY_NAME", Job.SupplyName
    SetAttribFromBox block, "SUPPLY_AVAILABLE_FLOW", Job.SupplyAvailableFlow
    SetAttribFromBox block, "SUPPLY_STATIC_PRESSURE", Job.SupplyStaticPressure
    SetAttribFromBox block, "SUPPLY_RESIDUAL_PRESSURE", Job.SupplyResidualPressure
    SetAttribFromBox block, "SUPPLY_ELEVATION", Job.SupplyElevation
    SetAttribFromBox block, "SUPPLY_MANIFOLD_PIPE_LENGTH", Job.SupplyManifoldPipeLength
    SetAttribFromBox block, "SUPPLY_MANIFOLD_PIPE_TYPE", Job.SupplyManifoldPipeType
    SetAttribFromBox block, "SUPPLY_MANIFOLD_PIPE_SIZE", Job.SupplyManifoldPipeSize
    SetAttribFromBox block, "SUPPLY_MANIFOLD_PIPE_INTERNAL_DIAMETER", Job.SupplyManifoldPipeInternalDiameter
    SetAttribFromBox block, "SUPPLY_MANIFOLD_FITTINGS_SUMMARY", Job.SupplyManifoldFittingsSummary
    SetAttribFromBox block, "SUPPLY_MANIFOLD_FITTINGS_EQUIV_LENGTH", Job.SupplyManifoldFittingsEquivLength
    SetAttribFromBox block, "SUPPLY_MANIFOLD_WATER_FLOW_PRESSURE_LOSS", Job.SupplyManifoldWaterFlowPressureLoss
    SetAttribFromBox block, "WATER_FLOW_MAKE_MODEL", Job.WaterFlowMakeModel
    SetAttribFromBox block, "DOMESTIC_FLOW_ADDED", Job.DomesticFlowAdded
    
    'SetAttribFromBox block, "STATIC_PRESSURE", Job.StaticPressure
    'SetAttribFromBox block, "RESIDUAL_PRESSURE", Job.ResidualPressure
    'SetAttribFromBox block, "AVAILABLE_FLOW", Job.AvailableFlow
    'SetAttribFromBox block, "METER_ELEVATION", Job.MeterElevation
    'SetAttribFromBox block, "METER_PIPE_LENGTH", Job.MeterPipeLength
    'SetAttribFromBox block, "METER_PIPE_INTERNAL_DIAMETER", Job.MeterPipeInternalDiameter
    
    CheckAddLayer "Job Data", Colors.vicMagenta, False, True
    block.layer = "Job Data"
    
    

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
Dim block As IntelliCAD.blockInsert
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
            Set block = entit
            If block.Name = "JobData" Then
            
              
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
    
                strJobName = GetAttrib(block, "JOB_NAME")
                strLeadNumber = GetAttrib(block, "LEAD_NUMBER")
                
                strSiteLocation = GetFirstAttrib(block, "SITE_LOCATION")
                strStaticPressure = GetFirstAttrib(block, "STATIC_PRESSURE")
                strResidualPressure = GetFirstAttrib(block, "RESIDUAL_PRESSURE")
                strAvailableFlow = GetFirstAttrib(block, "AVAILABLE_FLOW")
                strMeterElevation = GetFirstAttrib(block, "METER_ELEVATION")
                strMeterPipeLength = GetFirstAttrib(block, "METER_PIPE_LENGTH")
                strMeterPipeInternalDiameter = GetFirstAttrib(block, "METER_PIPE_INTERNAL_DIAMETER")
                strCalculatedByCompany = GetFirstAttrib(block, "CALCULATED_BY_COMPANY")
                
                Set Job = New Job
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

                Set Job.BlockReference = block
                blockCount = blockCount + 1
            End If
        End If
    Next i
    
    If blockCount = 0 Then Set Job = Nothing
    
End Sub
Public Function GetFirstAttrib(block As IntelliCAD.blockInsert, ParamArray tags())
Dim v As String
Dim tag As Variant

    For Each tag In tags
        v = GetAttrib(block, tag)
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
    Dim block As IntelliCAD.blockInsert
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
    
    Set block = Job.BlockReference
    
    strJobNameOld = GetAttrib(block, "JOB_NAME")
    strLeadNumberOld = GetAttrib(block, "LEAD_NUMBER")
    strSiteLocationOld = GetAttrib(block, "SITE_LOCATION")
    strStaticPressureOld = GetAttrib(block, "STATIC_PRESSURE")
    strResidualPressureOld = GetAttrib(block, "RESIDUAL_PRESSURE")
    strAvailableFlowOld = GetAttrib(block, "AVAILABLE_FLOW")
    strMeterElevationOld = GetAttrib(block, "METER_ELEVATION")
    strMeterPipeLengthOld = GetAttrib(block, "METER_PIPE_LENGTH")
    strMeterPipeInternalDiameterOld = GetAttrib(block, "METER_PIPE_INTERNAL_DIAMETER")
    strCalculatedByCompany = GetAttrib(block, "CALCULATED_BY_COMPANY")
    
    strJobNameNew = InputBox("Input Job Name", , strJobNameOld)
    If strJobNameNew = "" Then Exit Sub
    SetAttrib block, "JOB_NAME", strJobNameNew
    
    strLeadNumberNew = InputBox("Input Lead Number", , strLeadNumberOld)
    If strLeadNumberNew = "" Then Exit Sub
    SetAttrib block, "LEAD_NUMBER", strLeadNumberNew
    
    strSiteLocationNew = InputBox("Site Location", , strSiteLocationOld)
    SetAttrib block, "SITE_LOCATION", strSiteLocationNew
    
    strStaticPressureNew = InputBox("Static Pressure (psi)", , strStaticPressureOld)
    SetAttribFromBox block, "STATIC_PRESSURE", strStaticPressureNew
    
    strResidualPressureNew = InputBox("Residual Pressure (psi)", , strResidualPressureOld)
    SetAttribFromBox block, "RESIDUAL_PRESSURE", strResidualPressureNew
    
    strAvailableFlowNew = InputBox("Available Flow (gpm)", , strAvailableFlowOld)
    SetAttribFromBox block, "AVAILABLE_FLOW", strAvailableFlowNew
    
    strMeterElevationNew = InputBox("Elevation of meter (ft)", , strMeterElevationOld)
    SetAttribFromBox block, "METER_ELEVATION", strMeterElevationNew

    
    strMeterPipeLengthNew = InputBox("Length of pipe to meter (ft)", , strMeterPipeLengthOld)
    SetAttribFromBox block, "METER_PIPE_LENGTH", strMeterPipeLengthNew
    
    strMeterPipeInternalDiameterNew = InputBox("Internal diameter of pipe to meter (inches)", , strMeterPipeInternalDiameterOld)
    SetAttribFromBox block, "METER_PIPE_INTERNAL_DIAMETER", strMeterPipeInternalDiameterNew
    
    CheckAddLayer "Job Data", Colors.vicMagenta, False, True
    block.layer = "Job Data"
    
    

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

