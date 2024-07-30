VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ParticipantValuesRangeFiller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Preparation Document Templates.Values Ranges Filler"
Option Explicit
Implements IParticipantValuesRangeFiller

Private Type TParticipantValuesRangeFiller

    ValuesHandler As IMeasuredValuesHandler
    Participants As List_IParticipant_
    
End Type

Private this As TParticipantValuesRangeFiller

Public Property Get Self() As IParticipantValuesRangeFiller
    Set Self = Me
End Property

Public Property Get ValuesHandler() As IMeasuredValuesHandler
    Set ValuesHandler = this.ValuesHandler
End Property
Public Property Set ValuesHandler(ByVal RHS As IMeasuredValuesHandler)
    Set this.ValuesHandler = RHS
End Property

Public Property Get Participants() As List_IParticipant_
    Set Participants = this.Participants
End Property
Public Property Set Participants(ByVal RHS As List_IParticipant_)
    Set this.Participants = RHS
End Property
Public Function GetParticipantsWithValuesRanges() As List_IParticipant_

    Dim Result As List_IParticipant_
    Set Result = New List_IParticipant_

    Dim i As Long
    For i = 1 To Participants.Count

        Dim FilledParticipant As IComparisonDataParticipant
        Set FilledParticipant = GetFilledWithValuesRangesParticipant(Participants(i))

        Result.Add FilledParticipant

    Next
    
    Set GetParticipantsWithValuesRanges = Result

End Function
    Private Function GetFilledWithValuesRangesParticipant( _
            ByVal Participant As IComparisonDataParticipant) As IComparisonDataParticipant
        
        Dim ResultList As List_IParticipantResult_
        Set ResultList = New List_IParticipantResult_
        
        Dim i As Long
        For i = 1 To Participant.Results.Count
            
            Dim Result As IParticipantResult
            Set Result = GetFilledResult(Participant.Results(i))
            
            ResultList.Add Result
            
        Next
        
        Dim NewParticipant As ParticipantDTO
        Set NewParticipant = Participant
         
        Set NewParticipant.Results = ResultList
        Set GetFilledWithValuesRangesParticipant = NewParticipant
        
    End Function
        
        Private Function GetFilledResult( _
                ByVal SourceResult As IParticipantResult) As IParticipantResult
            
            Dim Result As ParticipantResultDTO
            Set Result = SourceResult
            
            Dim ValueNameKey As String
            ValueNameKey = Result.ValueName.Value & ". " & _
                           Result.ValueSubname.Value & ". " & _
                           Result.ValueUnit.Value
                    
            Result.ValueRange = ValuesHandler.GetMeasuredValueRange(ValueNameKey)
            
            Set GetFilledResult = Result
            
        End Function
'---------------------------------------------
Private Function IParticipantValuesRangeFiller_GetParticipantsWithValuesRanges() As List_IParticipant_
    Set IParticipantValuesRangeFiller_GetParticipantsWithValuesRanges = GetParticipantsWithValuesRanges
End Function


