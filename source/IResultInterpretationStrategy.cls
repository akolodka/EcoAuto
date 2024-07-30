VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IResultInterpretationStrategy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("eco_transferToWord.Core.Transfer Process.Result Interpretation Handler.Strategies")
'@Interface
Option Explicit

Public Function IsMatched(ByVal ParticipantResults As List_IParticipantResult_) As Boolean
End Function

Public Property Get BookmarkName() As String
End Property
