VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ResultInterpretationHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("eco_transferToWord.Core.Transfer Process.Result Interpretation Handler")
Option Explicit
Implements IResultInterpretationHandler

Private Type TResultInterpretationHandler

    Strategies As List_IResultInterpretationStra_
    Document As IWordDocument
    ParticipantResults As List_IParticipantResult_
    
    Progress As IProgressService
    
End Type

Private this As TResultInterpretationHandler

Public Property Get Self() As IResultInterpretationHandler
    Set Self = Me
End Property

Public Property Get Strategies() As List_IResultInterpretationStra_
    Set Strategies = this.Strategies
End Property
Public Property Set Strategies(ByVal RHS As List_IResultInterpretationStra_)
    Set this.Strategies = RHS
End Property

Public Property Get Document() As IWordDocument
    Set Document = this.Document
End Property
Public Property Set Document(ByVal RHS As IWordDocument)
    Set this.Document = RHS
End Property

Public Property Get ParticipantResults() As List_IParticipantResult_
    Set ParticipantResults = this.ParticipantResults
End Property
Public Property Set ParticipantResults(ByVal RHS As List_IParticipantResult_)
    Set this.ParticipantResults = RHS
End Property

Public Property Get Progress() As IProgressService
    Set Progress = this.Progress
End Property
Public Property Set Progress(ByVal RHS As IProgressService)
    Set this.Progress = RHS
End Property

Public Sub WriteInterpretation()
    
    Progress.Show _
        "Заполнение шаблона «" & Document.Name & "»", _
        "Заключение по результатам МСИ"
    
    Dim i As Long
    For i = 1 To Strategies.Count
    
        If (Strategies(i).IsMatched(ParticipantResults) = False) Then _
            RemoveInterpretation Strategies(i).BookmarkName, Document
    
    Next
    
End Sub
    Private Sub RemoveInterpretation( _
            ByVal BookmarkName As String, _
            ByVal Document As IEditableWordDocument)
        
        Dim TargetRange As IEditableWordRange
        Set TargetRange = Document.BookmarkedRange(BookmarkName)
        
        If (TargetRange Is Nothing = True) Then _
            Exit Sub
            
        TargetRange.ClearContent
        
    End Sub
'---------------------------------------------
Private Sub IResultInterpretationHandler_WriteInterpretation()
    WriteInterpretation
End Sub
