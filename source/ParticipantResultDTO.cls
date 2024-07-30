VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ParticipantResultDTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Initialization.Comparison Mapping.Model.Participant Results"
Option Explicit

Implements IParticipantResult

Implements ITableRowModel
Implements IMeasuredTableRowModel
Implements IReferenceTableRowModel
Implements IComparisonTableParticipantResu
Implements IComparisonResultHeaderModel
Implements IComparisonChartModel
Implements IResultEstimationTableRowModel
Implements IResultInterpretationModel

Private Type TParticipantResultDTO

    ValueName As IKeyValuePair
    ValueSubname As IKeyValuePair
    ValueUnit As IKeyValuePair
    
    ValueRange As String
    
    ParticipantValue As String
    ParticipantUncertainty As String
    
    ReferenceValue As String
    ReferenceUncertainty As String
    
    EnCriteria As String
    
End Type

Private this As TParticipantResultDTO

Public Property Get ValueName() As IKeyValuePair
    Set ValueName = this.ValueName
End Property
Public Property Set ValueName(ByVal RHS As IKeyValuePair)
    Set this.ValueName = RHS
End Property

Public Property Get ValueSubname() As IKeyValuePair
    Set ValueSubname = this.ValueSubname
End Property
Public Property Set ValueSubname(ByVal RHS As IKeyValuePair)
    Set this.ValueSubname = RHS
End Property

Public Property Get ValueUnit() As IKeyValuePair
    Set ValueUnit = this.ValueUnit
End Property
Public Property Set ValueUnit(ByVal RHS As IKeyValuePair)
    Set this.ValueUnit = RHS
End Property

Public Property Get ValueRange() As String
    ValueRange = this.ValueRange
End Property
Public Property Let ValueRange(ByVal RHS As String)
    this.ValueRange = RHS
End Property

Public Property Get ParticipantValue() As String
    ParticipantValue = this.ParticipantValue
End Property
Public Property Let ParticipantValue(ByVal RHS As String)
    this.ParticipantValue = RHS
End Property

Public Property Get ParticipantUncertainty() As String
    ParticipantUncertainty = this.ParticipantUncertainty
End Property
Public Property Let ParticipantUncertainty(ByVal RHS As String)
    this.ParticipantUncertainty = RHS
End Property

Public Property Get ReferenceValue() As String
    ReferenceValue = this.ReferenceValue
End Property
Public Property Let ReferenceValue(ByVal RHS As String)
    this.ReferenceValue = RHS
End Property

Public Property Get ReferenceUncertainty() As String
    ReferenceUncertainty = this.ReferenceUncertainty
End Property
Public Property Let ReferenceUncertainty(ByVal RHS As String)
    this.ReferenceUncertainty = RHS
End Property

Public Property Get EnCriteria() As String
    EnCriteria = this.EnCriteria
End Property
Public Property Let EnCriteria(ByVal RHS As String)
    this.EnCriteria = RHS
End Property

Public Property Get Estimation() As String
    
    Dim Result As String
    Result = "неудовлетворительно"
    
    If IsNumeric(EnCriteria) = False Then _
        Exit Property
    
    If CDbl(EnCriteria) < 1 Then _
        Result = "удовлетворительно"
        
    Estimation = Result
    
End Property

'---------------------------------------------
Private Property Get IMeasuredTableRowModel_Name() As String
    IMeasuredTableRowModel_Name = ValueName.Value & ", " & ValueUnit.Value
End Property
Private Property Get IMeasuredTableRowModel_Subname() As String
    IMeasuredTableRowModel_Subname = ValueSubname.Value
End Property
Private Property Get IMeasuredTableRowModel_Range() As String
    IMeasuredTableRowModel_Range = ValueRange
End Property
'---------------------------------------------
Private Property Get IReferenceTableRowModel_Name() As String
    IReferenceTableRowModel_Name = ValueName.Value & ", " & ValueUnit.Value
End Property
Private Property Get IReferenceTableRowModel_Subname() As String
    IReferenceTableRowModel_Subname = ValueSubname.Value
End Property
Private Property Get IReferenceTableRowModel_Value() As String
    IReferenceTableRowModel_Value = ReferenceValue
End Property
Private Property Get IReferenceTableRowModel_Uncertainty() As String
    IReferenceTableRowModel_Uncertainty = ReferenceUncertainty
End Property
'---------------------------------------------
Private Property Get IComparisonTableParticipantResu_Value() As String
    IComparisonTableParticipantResu_Value = ParticipantValue
End Property
Private Property Get IComparisonTableParticipantResu_Uncertainty() As String
    IComparisonTableParticipantResu_Uncertainty = ParticipantUncertainty
End Property
Private Property Get IComparisonTableParticipantResu_EnCriteria() As String
    IComparisonTableParticipantResu_EnCriteria = EnCriteria
End Property
'---------------------------------------------
Private Property Get IComparisonResultHeaderModel_Name() As IKeyValuePair
    Set IComparisonResultHeaderModel_Name = ValueName
End Property
Private Property Get IComparisonResultHeaderModel_SubName() As IKeyValuePair
    Set IComparisonResultHeaderModel_SubName = ValueSubname
End Property
Private Property Get IComparisonResultHeaderModel_Unit() As IKeyValuePair
    Set IComparisonResultHeaderModel_Unit = ValueUnit
End Property
'---------------------------------------------
Private Property Get IComparisonChartModel_ValueName() As String
    IComparisonChartModel_ValueName = ValueName.Value & ", " & ValueUnit.Value & ". " & ValueSubname.Value
End Property
Private Property Get IComparisonChartModel_ParticipantValue() As String
    IComparisonChartModel_ParticipantValue = ParticipantValue
End Property
Private Property Get IComparisonChartModel_ParticipantUncertainty() As String
    IComparisonChartModel_ParticipantUncertainty = ParticipantUncertainty
End Property
Private Property Get IComparisonChartModel_ReferenceValue() As String
    IComparisonChartModel_ReferenceValue = ReferenceValue
End Property
Private Property Get IComparisonChartModel_ReferenceUncertainty() As String
    IComparisonChartModel_ReferenceUncertainty = ReferenceUncertainty
End Property
'---------------------------------------------
Private Property Get IResultEstimationTableRowModel_Name() As String
    IResultEstimationTableRowModel_Name = ValueName.Value & ", " & ValueUnit.Value
End Property
Private Property Get IResultEstimationTableRowModel_SubName() As String
    IResultEstimationTableRowModel_SubName = ValueSubname.Value
End Property
Private Property Get IResultEstimationTableRowModel_EnCriteria() As String
    IResultEstimationTableRowModel_EnCriteria = EnCriteria
End Property
Private Property Get IResultEstimationTableRowModel_Estimation() As String
    IResultEstimationTableRowModel_Estimation = Estimation
End Property
'---------------------------------------------
Private Property Get IResultInterpretationModel_Estimation() As String
    IResultInterpretationModel_Estimation = Estimation
End Property
