VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ParticipantTablesDecorator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Comparison Results Decoration.Participant Tables Decoration"
Option Explicit
Implements IParticipantTablesDecorator

Private Type TInputConditionsDecorator
    
    Strategy As ITableDecorationStrategy
    Document As ISavableWordDocument
    ParticipantResults As List_IParticipantResult_
    
End Type

Private this As TInputConditionsDecorator

Public Property Get Self() As IParticipantTablesDecorator
    Set Self = Me
End Property

Public Property Get Strategy() As ITableDecorationStrategy
    Set Strategy = this.Strategy
End Property
Public Property Set Strategy(ByVal RHS As ITableDecorationStrategy)
    Set this.Strategy = RHS
End Property

Public Property Get Document() As ISavableWordDocument
    Set Document = this.Document
End Property
Public Property Set Document(ByVal RHS As ISavableWordDocument)
    Set this.Document = RHS
End Property

Public Property Get ParticipantResults() As List_IParticipantResult_
    Set ParticipantResults = this.ParticipantResults
End Property
Public Property Set ParticipantResults(ByVal RHS As List_IParticipantResult_)
    Set this.ParticipantResults = RHS
End Property

Private Sub ExecuteFilling()
    
    If (ParticipantResults.Count = vbEmpty) Then _
        Exit Sub
    
    Dim Table As IWordTable
    Set Table = GetWorkTable()
    
    Dim i As Long
    For i = 1 To ParticipantResults.Count
        
        If (IsDataObtained(ParticipantResults(i)) = True) Then _
            AppendRow Table, ParticipantResults(i)
        
    Next
    
    ApplyTableDecoration Table
    
End Sub
    Private Function GetWorkTable() As IWordTable
        
        Dim Handler As IDecorationTableHandler
        Set Handler = DecorationTableHandler.Create(Document)
            
        Dim Result As IWordTable
        Set Result = Handler.GetTable()
            
        Set GetWorkTable = Result
    
    End Function
    
    Private Function IsDataObtained( _
            ByVal Model As IComparisonTableParticipantResu) As Boolean
            
        IsDataObtained = (Model.Value <> vbNullString) And (Model.Uncertainty <> vbNullString)
        
    End Function
    Private Sub AppendRow( _
        ByVal Table As IWordTable, _
        ByVal Model As ITableRowModel)
        
        Dim Decorator As ITableRowAppendStrategy
        Set Decorator = Strategy
        
        Decorator.AppendRow Table, Model
        
    End Sub
    Private Sub ApplyTableDecoration( _
            ByVal Table As IWordTable)
            
        Dim Decorator As ITableDecorationStrategy
        Set Decorator = Strategy
        
        Decorator.ApplyTableDecoration Table
    
    End Sub
Private Sub SaveChanges()

    If (Document Is Nothing = True) Then _
        Exit Sub
        
    Document.Save
End Sub
'---------------------------------------------
Private Sub IParticipantTablesDecorator_ExecuteFilling()
    ExecuteFilling
End Sub
Private Sub IParticipantTablesDecorator_SaveChanges()
    SaveChanges
End Sub


