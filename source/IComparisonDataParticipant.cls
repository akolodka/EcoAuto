VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IComparisonDataParticipant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Initialization.Comparison Mapping.Model.Participant.Contracts"
'@Interface
Option Explicit

Public Property Get Number() As IKeyValuePair
End Property

Public Property Get Tour() As IKeyValuePair
End Property

Public Property Get TypeP() As IKeyValuePair
End Property

Public Property Get Organisation() As IKeyValuePair
End Property

Public Property Get Results() As List_IParticipantResult_
End Property

