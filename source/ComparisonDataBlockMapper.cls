VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ComparisonDataBlockMapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("eco_transferToWord.Core.Initialization.Comparison Mapping.Block Data Mapper")
'@PredeclaredId
Option Explicit
Implements IComparisonDataBlockMapper

Const BlockInitialColumn As Long = 7
Const BlockStep As Long = 6

Const ValueNameKeyColumnShift As Long = 0
Const ValueNameValueColumnShift As Long = 1

Const ValueSubnameKeyColumnShift As Long = 2
Const ValueSubnameValueColumnShift As Long = 3

Const ValueUnitKeyColumnShift As Long = 4
Const ValueUnitValueColumnShift As Long = 5

Const ParticipantValueColumnShift As Long = 0
Const ParticipantUncertaintyColumnShift As Long = 1

Const ReferenceValueColumnShift As Long = 3
Const ReferencetUncertaintyColumnShift As Long = 4

Const EnCriteriaColumnShift As Long = 5

Private Type TComparisonDataBlockMapper
    
    KeysRow As IComparisonDataRow
    ParticipantRow As IComparisonDataRow
    
End Type

Private this As TComparisonDataBlockMapper

Public Function Create( _
       ByVal KeysRow As IComparisonDataRow, _
       ByVal ParticipantRow As IComparisonDataRow) As IComparisonDataBlockMapper
    
    With New ComparisonDataBlockMapper
        
        Set .KeysRow = KeysRow
        Set .ParticipantRow = ParticipantRow
        
        Set Create = .Self
    
    End With

End Function

Public Property Get KeysRow() As IComparisonDataRow
    Set KeysRow = this.KeysRow
End Property
Public Property Set KeysRow(ByVal RHS As IComparisonDataRow)
    Set this.KeysRow = RHS
End Property

Public Property Get ParticipantRow() As IComparisonDataRow
    Set ParticipantRow = this.ParticipantRow
End Property
Public Property Set ParticipantRow(ByVal RHS As IComparisonDataRow)
    Set this.ParticipantRow = RHS
End Property

Public Property Get Self() As IComparisonDataBlockMapper
    Set Self = Me
End Property

Public Function GetResults() As List_IParticipantResult_
     
    Dim Results As List_IParticipantResult_
    Set Results = New List_IParticipantResult_
        
    Dim Index As Long
    Index = BlockInitialColumn
        
    Do While Index < KeysRow.ColumnsCount
        
        Dim Model As IParticipantResult
        Set Model = GetComparisonResult(Index)
        
        Results.Add Model
        Index = Index + BlockStep
            
    Loop
    
    Set GetResults = Results
        
End Function
    Private Function IsComparisonResultObtained( _
            ByVal Column As Long) As Boolean
            
        Dim Value As String
        Value = ParticipantRow.GetValue(Column + ParticipantValueColumnShift)
        
        Dim Uncertainty As String
        Uncertainty = ParticipantRow.GetValue(Column + ParticipantUncertaintyColumnShift)
        
        IsComparisonResultObtained = (Value <> vbNullString) And (Uncertainty <> vbNullString)
        
    End Function
    Private Function GetComparisonResult( _
            ByVal InitialColumn As Long) As IParticipantResult
                    
        Dim Result As ParticipantResultDTO
        Set Result = New ParticipantResultDTO
                    
        Set Result.ValueName = KeyValuePair.Create( _
                                    KeysRow.GetValue(InitialColumn + ValueNameKeyColumnShift), _
                                    KeysRow.GetValue(InitialColumn + ValueNameValueColumnShift))
                        
        Set Result.ValueSubname = KeyValuePair.Create( _
                                    KeysRow.GetValue(InitialColumn + ValueSubnameKeyColumnShift), _
                                    KeysRow.GetValue(InitialColumn + ValueSubnameValueColumnShift))
                        
        Set Result.ValueUnit = KeyValuePair.Create( _
                                    KeysRow.GetValue(InitialColumn + ValueUnitKeyColumnShift), _
                                    KeysRow.GetValue(InitialColumn + ValueUnitValueColumnShift))
                                   
        Result.ValueRange = "..."
                                
        Dim Value As String
        Value = ParticipantRow.GetValue(InitialColumn + ParticipantValueColumnShift)
                     
        Dim Uncertainty As String
        Uncertainty = ParticipantRow.GetValue(InitialColumn + ParticipantUncertaintyColumnShift)
                                       
        Dim Rounder As IValueRounder
        Set Rounder = ValueRounder.Create(Value, Uncertainty)
                                        
        Result.ParticipantValue = Rounder.GetValue
        Result.ParticipantUncertainty = Rounder.GetUncertainty
        
        Value = ParticipantRow.GetValue(InitialColumn + ReferenceValueColumnShift)
        Uncertainty = ParticipantRow.GetValue(InitialColumn + ReferencetUncertaintyColumnShift)
          
        Set Rounder = ValueRounder.Create(Value, Uncertainty)
        
        Result.ReferenceValue = Rounder.GetValue
        Result.ReferenceUncertainty = Rounder.GetUncertainty
        
        Value = ParticipantRow.GetValue(InitialColumn + EnCriteriaColumnShift)
        Set Rounder = ValueRounder.Create(Value)
        
        Result.EnCriteria = Rounder.GetRoundedToHundredths
              
        Set GetComparisonResult = Result
                    
    End Function
'---------------------------------------------
Private Function IComparisonDataBlockMapper_GetComparisonResults() As List_IParticipantResult_
    Set IComparisonDataBlockMapper_GetComparisonResults = GetResults
End Function


