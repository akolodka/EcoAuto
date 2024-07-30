VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ComparisonDataParticipantMapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Initialization.Comparison Mapping.Participant Data Mapper"
'@PredeclaredId
Option Explicit
Implements IComparisonDataParticipantMappe

Const NumberColumn As Long = 2
Const TourColumn As Long = 3
Const TypeColumn As Long = 5
Const OrganisationColumn As Long = 6

Private Type TComparisonDataRowMapper

    KeysRow As IComparisonDataRow
    ParticipantRow As IComparisonDataRow
    
End Type

Private this As TComparisonDataRowMapper

Public Function Create( _
       ByVal KeysRow As IComparisonDataRow, _
       ByVal ParticipantRow As IComparisonDataRow) As IComparisonDataParticipantMappe
    
    With New ComparisonDataParticipantMapper
        
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

Public Property Get Self() As IComparisonDataParticipantMappe
    Set Self = Me
End Property

Public Function GetParticipant() As IParticipant
    
    Dim Participant As ParticipantDTO
    Set Participant = New ParticipantDTO
    
    Set Participant.Number = KeyValuePair.Create( _
                             KeysRow.GetValue(NumberColumn), _
                             ParticipantRow.GetValue(NumberColumn))
    
    Set Participant.Tour = KeyValuePair.Create( _
                           KeysRow.GetValue(TourColumn), _
                           ParticipantRow.GetValue(TourColumn))
    
    Set Participant.TypeP = KeyValuePair.Create( _
                                 KeysRow.GetValue(TypeColumn), _
                                 ParticipantRow.GetValue(TypeColumn))
    
    Set Participant.Organisation = KeyValuePair.Create( _
                                   KeysRow.GetValue(OrganisationColumn), _
                                   ParticipantRow.GetValue(OrganisationColumn))
                              
    Set GetParticipant = Participant
        
End Function
'---------------------------------------------
Private Function IComparisonDataParticipantMappe_GetParticipant() As IParticipant
    Set IComparisonDataParticipantMappe_GetParticipant = GetParticipant
End Function


