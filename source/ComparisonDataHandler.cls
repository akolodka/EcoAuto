VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ComparisonDataHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("eco_transferToWord.Core.Initialization.Comparison Mapping")
'@PredeclaredId
Option Explicit
Implements IComparisonDataHandler

Private Type TComparisonDataHandler
    
    Participants As List_IParticipant_
    Reader As IComparisonDataRowsReader
    
End Type

Private this As TComparisonDataHandler

Public Function Create( _
       ByVal Reader As IComparisonDataRowsReader) As IComparisonDataHandler
    
    With New ComparisonDataHandler
        
        Set .Reader = Reader
        .Initialize
        
        Set Create = .Self
        
    End With
    
End Function

Public Property Get Reader() As IComparisonDataRowsReader
    Set Reader = this.Reader
End Property
Public Property Set Reader(ByVal RHS As IComparisonDataRowsReader)
    Set this.Reader = RHS
End Property

Public Property Get Participants() As List_IParticipant_
    Set Participants = this.Participants
End Property
Public Property Set Participants(ByVal RHS As List_IParticipant_)
    Set this.Participants = RHS
End Property

Public Property Get Self() As IComparisonDataHandler
    Set Self = Me
End Property

Public Sub Initialize()
    
    Set Participants = New List_IParticipant_
    
    If (Reader Is Nothing = True) Then _
        Exit Sub
    
    Dim DataRows As List_IExcelRange_
    Set DataRows = Reader.GetComparisonDataRows()
    
    If (DataRows.Count = vbEmpty) Then _
        Exit Sub
    
    Dim i As Long
    For i = 1 To DataRows.Count
    
        Dim Model As IParticipant
        Set Model = GetParticipant(DataRows(i))
        
        If (IsCorrect(Model) = True) Then _
           Participants.Add Model
        
    Next
    
End Sub
    Private Function GetParticipant( _
            ByVal DataRow As IExcelRange) As IParticipant
        
        Dim BaseDataMapper As IComparisonDataParticipantMappe
        Set BaseDataMapper = ComparisonDataParticipantMapper.Create(Reader.GetKeysRow, DataRow)

        Dim Participant As ParticipantDTO
        Set Participant = BaseDataMapper.GetParticipant()
        
        Dim BlockMapper As IComparisonDataBlockMapper
        Set BlockMapper = ComparisonDataBlockMapper.Create(Reader.GetKeysRow, DataRow)
        
        Set Participant.Results = BlockMapper.GetComparisonResults()
        
        Set GetParticipant = Participant
        
    End Function
    Private Function IsCorrect( _
            ByVal Participant As IComparisonDataParticipant) As Boolean
                
        If (Participant.Number.Value = vbNullString) Then _
           Exit Function
                    
        If (Participant.Organisation.Value = vbNullString) Then _
           Exit Function
                
        IsCorrect = True
    
    End Function
'---------------------------------------------
Private Function IComparisonDataHandler_GetParticipants() As List_IParticipant_
    Set IComparisonDataHandler_GetParticipants = Participants
End Function


