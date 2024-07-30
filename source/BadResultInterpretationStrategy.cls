VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BadResultInterpretationStrategy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("eco_transferToWord.Core.Transfer Process.Result Interpretation Handler.Strategies")
Option Explicit
Implements IResultInterpretationStrategy

Const EstimationCriteria As String = "неудовлетворительно"
Const BookmarkName As String = "BadInterpretation"

Public Function IsMatched( _
       ByVal ParticipantResults As List_IParticipantResult_) As Boolean
    
    Dim i As Long
    For i = 1 To ParticipantResults.Count
        
        Dim Model As IResultInterpretationModel
        Set Model = ParticipantResults(i)
        
        If Model.Estimation = EstimationCriteria Then _
            IsMatched = True: _
            Exit Function
        
    Next
    
End Function
Private Function IResultInterpretationStrategy_IsMatched(ByVal ParticipantResults As List_IParticipantResult_) As Boolean
    IResultInterpretationStrategy_IsMatched = IsMatched(ParticipantResults)
End Function
Private Property Get IResultInterpretationStrategy_BookmarkName() As String
    IResultInterpretationStrategy_BookmarkName = BookmarkName
End Property

