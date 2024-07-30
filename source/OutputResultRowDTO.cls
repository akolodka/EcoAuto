VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OutputResultRowDTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Comparison Results Decoration.Comparison Block Decoration.Model"
Option Explicit

Implements ITableRowModel
Implements IOutputResultsRowModel
Implements IComparisonTableRowModel
Implements IComparisonChartSheetRowModel
Implements IComparisonChartIdentifyModel
Implements IComparisonChartSheetHeaderMode

Private Type TComparisonTableRowDTO
    
    Number As String
    Name As String
    
    ParticipantValue As String
    ParticipantUncertainty As String
    
    ReferenceValue As String
    ReferenceUncertainty As String
    
    EnCriteria As String
    ValueName As String
    
    BlockIndex As Long
    
End Type

Private this As TComparisonTableRowDTO

Public Property Get Self() As IComparisonTableRowModel
    Set Self = Me
End Property

Public Property Get Number() As String
    Number = this.Number
End Property
Public Property Let Number(ByVal RHS As String)
    this.Number = RHS
End Property

Public Property Get Name() As String
    Name = this.Name
End Property
Public Property Let Name(ByVal RHS As String)
    this.Name = RHS
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

Public Property Get ValueName() As String
    ValueName = this.ValueName
End Property
Public Property Let ValueName(ByVal RHS As String)
    this.ValueName = RHS
End Property

Public Property Get BlockIndex() As Long
    BlockIndex = this.BlockIndex
End Property
Public Property Let BlockIndex(ByVal RHS As Long)
    this.BlockIndex = RHS
End Property

'---------------------------------------------
Private Property Get IComparisonTableRowModel_Number() As String
    IComparisonTableRowModel_Number = Number
End Property
Private Property Get IComparisonTableRowModel_Name() As String
    IComparisonTableRowModel_Name = Name
End Property
Private Property Get IComparisonTableRowModel_Value() As String
    IComparisonTableRowModel_Value = ParticipantValue
End Property
Private Property Get IComparisonTableRowModel_Uncertainty() As String
    IComparisonTableRowModel_Uncertainty = ParticipantUncertainty
End Property
Private Property Get IComparisonTableRowModel_EnCriteria() As String
    IComparisonTableRowModel_EnCriteria = EnCriteria
End Property
'---------------------------------------------
Private Property Get IOutputResultsRowModel_ParticipantNumber() As String
    IOutputResultsRowModel_ParticipantNumber = Number
End Property

Private Property Get IOutputResultsRowModel_BlockIndex() As Long
    IOutputResultsRowModel_BlockIndex = BlockIndex
End Property
'---------------------------------------------
Private Property Get IComparisonChartSheetRowModel_ParticipantName() As String
    IComparisonChartSheetRowModel_ParticipantName = Name
End Property
Private Property Get IComparisonChartSheetRowModel_Value() As String
    IComparisonChartSheetRowModel_Value = ParticipantValue
End Property
Private Property Get IComparisonChartSheetRowModel_Uncertainty() As String
    IComparisonChartSheetRowModel_Uncertainty = ParticipantUncertainty
End Property
'---------------------------------------------
Private Property Get IComparisonChartIdentifyModel_ParticipantName() As String
    IComparisonChartIdentifyModel_ParticipantName = Name
End Property
Private Property Get IComparisonChartIdentifyModel_BlockIndex() As Long
    IComparisonChartIdentifyModel_BlockIndex = BlockIndex
End Property
'---------------------------------------------
Private Property Get IComparisonChartSheetHeaderMode_ReferenceValue() As String
    IComparisonChartSheetHeaderMode_ReferenceValue = ReferenceValue
End Property
Private Property Get IComparisonChartSheetHeaderMode_ReferenceUncertainty() As String
    IComparisonChartSheetHeaderMode_ReferenceUncertainty = ReferenceUncertainty
End Property
Private Property Get IComparisonChartSheetHeaderMode_ValueName() As String
    IComparisonChartSheetHeaderMode_ValueName = ValueName
End Property
