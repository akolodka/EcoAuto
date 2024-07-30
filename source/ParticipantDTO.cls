VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ParticipantDTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Initialization.Comparison Mapping.Model.Participant"
Option Explicit

Implements IParticipant
Implements IComparisonDataParticipant
Implements ITransferMenuParticipant
Implements IComparisonTableParticipant
Implements IResultsObtainedParticipant

Private Type TParticipantDTO

    Number As IKeyValuePair
    Tour As IKeyValuePair
    
    Type As IKeyValuePair
    Organisation As IKeyValuePair
    
    Results As List_IParticipantResult_
    
End Type

Private this As TParticipantDTO

Public Property Get Number() As IKeyValuePair
    Set Number = this.Number
End Property
Public Property Set Number(ByVal RHS As IKeyValuePair)
    Set this.Number = RHS
End Property

Public Property Get Tour() As IKeyValuePair
    Set Tour = this.Tour
End Property
Public Property Set Tour(ByVal RHS As IKeyValuePair)
    Set this.Tour = RHS
End Property

Public Property Get TypeP() As IKeyValuePair
    Set TypeP = this.Type
End Property
Public Property Set TypeP(ByVal RHS As IKeyValuePair)
    Set this.Type = RHS
End Property

Public Property Get Organisation() As IKeyValuePair
    Set Organisation = this.Organisation
End Property
Public Property Set Organisation(ByVal RHS As IKeyValuePair)
    Set this.Organisation = RHS
End Property

Public Property Get Results() As List_IParticipantResult_
    Set Results = this.Results
End Property
Public Property Set Results(ByVal RHS As List_IParticipantResult_)
    Set this.Results = RHS
End Property
'---------------------------------------------
Private Property Get IComparisonDataParticipant_Number() As IKeyValuePair
    Set IComparisonDataParticipant_Number = Number
End Property
Private Property Get IComparisonDataParticipant_Tour() As IKeyValuePair
    Set IComparisonDataParticipant_Tour = Tour
End Property
Private Property Get IComparisonDataParticipant_TypeP() As IKeyValuePair
    Set IComparisonDataParticipant_TypeP = TypeP
End Property
Private Property Get IComparisonDataParticipant_Organisation() As IKeyValuePair
    Set IComparisonDataParticipant_Organisation = Organisation
End Property
Private Property Get IComparisonDataParticipant_Results() As List_IParticipantResult_
    Set IComparisonDataParticipant_Results = Results
End Property
'---------------------------------------------
Private Property Get ITransferMenuParticipant_Number() As String
    ITransferMenuParticipant_Number = Number.Value
End Property
Private Property Get ITransferMenuParticipant_Tour() As String
    ITransferMenuParticipant_Tour = Tour.Value
End Property
Private Property Get ITransferMenuParticipant_TypeP() As String
    ITransferMenuParticipant_TypeP = TypeP.Value
End Property
Private Property Get ITransferMenuParticipant_Organisation() As String
    ITransferMenuParticipant_Organisation = Organisation.Value
End Property
'---------------------------------------------
Private Property Get IComparisonTableParticipant_Number() As String
    IComparisonTableParticipant_Number = Number.Value
End Property
Private Property Get IComparisonTableParticipant_Name() As String
    IComparisonTableParticipant_Name = Organisation.Value
End Property
Private Property Get IComparisonTableParticipant_Results() As List_IParticipantResult_
    Set IComparisonTableParticipant_Results = Results
End Property
'---------------------------------------------
Private Property Get IResultsObtainedParticipant_Results() As List_IParticipantResult_
    Set IResultsObtainedParticipant_Results = Results
End Property
