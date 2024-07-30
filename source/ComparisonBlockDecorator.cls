VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ComparisonBlockDecorator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Comparison Results Decoration.Comparison Block Decoration"
Option Explicit
Implements IComparisonBlockDecorator

Private Type TComparisonBlockDecorator

    Document As ISavableWordDocument
    
    Multiplier As IComparisonBlockMultiplier
    Writer As IComparisonHeaderWriter
    
    Participants As List_IParticipant_
    Decorator As IComparisonBlockDecorationStrat
    
End Type

Private this As TComparisonBlockDecorator

Public Property Get Self() As IComparisonBlockDecorator
    Set Self = Me
End Property

Public Property Get Document() As ISavableWordDocument
    Set Document = this.Document
End Property
Public Property Set Document(ByVal RHS As ISavableWordDocument)
    Set this.Document = RHS
End Property

Public Property Get Multiplier() As IComparisonBlockMultiplier
    Set Multiplier = this.Multiplier
End Property
Public Property Set Multiplier(ByVal RHS As IComparisonBlockMultiplier)
    Set this.Multiplier = RHS
End Property

Public Property Get Writer() As IComparisonHeaderWriter
    Set Writer = this.Writer
End Property
Public Property Set Writer(ByVal RHS As IComparisonHeaderWriter)
    Set this.Writer = RHS
End Property

Public Property Get Participants() As List_IParticipant_
    Set Participants = this.Participants
End Property
Public Property Set Participants(ByVal RHS As List_IParticipant_)
    Set this.Participants = RHS
End Property

Public Property Get Decorator() As IComparisonBlockDecorationStrat
    Set Decorator = this.Decorator
End Property
Public Property Set Decorator(ByVal RHS As IComparisonBlockDecorationStrat)
    Set this.Decorator = RHS
End Property

Public Sub ExecuteFilling( _
       ByVal ParticipantNumber As String, _
       ByVal ComparisonBlockIndex As Long)
    
    If (IsParticipantResultObtained(ParticipantNumber, ComparisonBlockIndex) = False) Then _
        Exit Sub
        
    If (IsParticipantResultTheOnlyOne(ParticipantNumber, ComparisonBlockIndex) = True) Then _
        Exit Sub
    
    Multiplier.ExecuteMultiplication
    '---------------------------------------------
    Dim Header As IComparisonResultHeaderModel
    Set Header = GetHeaderModel(ParticipantNumber, ComparisonBlockIndex)

    Writer.FillBlockHeaders Header
    '---------------------------------------------
    Dim Model As List_IOutputResultRowModel_
    Set Model = New List_IOutputResultRowModel_

    Dim i As Long
    For i = 1 To Participants.Count

        Dim RowData As IOutputResultsRowModel
        Set RowData = GetOutputModel(Participants(i), ComparisonBlockIndex)
        
        If (RowData Is Nothing = False) Then _
            Model.Add RowData

    Next

    Decorator.FillBlock Model, ParticipantNumber

End Sub
    Private Function IsParticipantResultObtained( _
            ByVal ParticipantNumber As String, _
            ByVal BlockIndex As Long) As Boolean
    
        Dim i As Long
        For i = 1 To Participants.Count
            
            Dim Number As String
            Number = GetNumber(Participants(i))
        
            If (Number = ParticipantNumber) Then
                
                Dim Results As List_IParticipantResult_
                Set Results = GetResults(Participants(i))
                
                IsParticipantResultObtained = (IsDataObtained(Results(BlockIndex)) = True)
                
            End If
            
        Next
    
    End Function
        Private Function GetNumber( _
                ByVal Model As IComparisonDataParticipant) As String
            
            GetNumber = Model.Number.Value
            
        End Function
        Private Function GetResults( _
                ByVal Model As IResultsObtainedParticipant) As List_IParticipantResult_
            
            Set GetResults = Model.Results
            
        End Function
        Private Function IsDataObtained( _
            ByVal Model As IComparisonTableParticipantResu) As Boolean

            IsDataObtained = (Model.Value <> vbNullString) And (Model.Uncertainty <> vbNullString)

        End Function
    
    Private Function IsParticipantResultTheOnlyOne( _
            ByVal ParticipantNumber As String, _
            ByVal BlockIndex As Long) As Boolean
        
        Dim FilledValuesCount As Long

        Dim i As Long
        For i = 1 To Participants.Count
        
            Dim Results As List_IParticipantResult_
            Set Results = GetResults(Participants(i))
            
            Dim Model As IComparisonTableParticipantResu
            Set Model = Results(BlockIndex)
            
            If Model.Value <> vbNullString Then _
                FilledValuesCount = FilledValuesCount + 1
            
        Next
        
        IsParticipantResultTheOnlyOne = (FilledValuesCount = 1)
        
    End Function
    
    Private Function GetHeaderModel( _
            ByVal ParticipantNumber As String, _
            ByVal BlockIndex As Long) As IComparisonResultHeaderModel
        
        Dim i As Long
        For i = 1 To Participants.Count
            
            Dim Number As String
            Number = GetNumber(Participants(i))
            
            Dim Results As List_IParticipantResult_
            Set Results = GetResults(Participants(i))
            
            If (Number = ParticipantNumber) Then _
                Set GetHeaderModel = Results(BlockIndex)
            
        Next

    End Function
    Private Function GetOutputModel( _
            ByVal Participant As IComparisonTableParticipant, _
            ByVal BlockIndex As Long) As IOutputResultsRowModel
        
        Dim Model As IComparisonTableParticipantResu
        Set Model = Participant.Results(BlockIndex)
        
        If (Model.Value = vbNullString) Or (Model.Uncertainty = vbNullString) Then _
            Exit Function
        
        With New OutputResultRowDTO
                        
            .Name = Participant.Name
            .Number = Participant.Number
            
            .ParticipantValue = Model.Value
            .ParticipantUncertainty = Model.Uncertainty
            
            .ReferenceValue = GetReferenceValue(Model)
            .ReferenceUncertainty = GetReferenceUncertainty(Model)

            .EnCriteria = Model.EnCriteria
            
            .ValueName = GetValueName(Model)
            .BlockIndex = BlockIndex

            Set GetOutputModel = .Self

        End With
    
    End Function
        Private Function GetReferenceValue(ByVal Model As IComparisonChartModel) As String
            GetReferenceValue = Model.ReferenceValue
        End Function
        Private Function GetReferenceUncertainty(ByVal Model As IComparisonChartModel) As String
            GetReferenceUncertainty = Model.ReferenceUncertainty
        End Function
        Private Function GetValueName(ByVal Model As IComparisonChartModel) As String
            GetValueName = Model.ValueName
        End Function
        
Private Sub Dispose()
    Multiplier.Dispose
End Sub
Private Sub SaveChanges()

    If (Document Is Nothing = True) Then _
        Exit Sub
            
    Document.Save
    
End Sub
'---------------------------------------------
Private Sub IComparisonBlockDecorator_ExecuteFilling(ByVal ParticipantNumber As String, ByVal ComparisonBlockIndex As Long)
    ExecuteFilling ParticipantNumber, ComparisonBlockIndex
End Sub
Private Sub IComparisonBlockDecorator_Dispose()
    Dispose
End Sub
Private Sub IComparisonBlockDecorator_SaveChanges()
    SaveChanges
End Sub


