VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ReferenceValuesValidator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Reference Values Validator"
'@PredeclaredId
Option Explicit
Implements IReferenceValuesValidator

Private Type TReferenceValuesValidator

    Participants As List_IParticipant_
    Worksheet As ICellFindableExcelWorksheet
    
    IsReferenceDataUnique As Boolean
    ValidationSummary As List_IComparisonBlockValidMode_

End Type

Private this As TReferenceValuesValidator

Public Function Create( _
       ByVal Initial As IReferenceValuesValidatorInitia) As IReferenceValuesValidator
    
    With New ReferenceValuesValidator
        
        Set .Participants = Initial.Participants
        Set .Worksheet = Initial.SourceWorksheet
        
        .Initialize
        
        Set Create = .Self
        
    End With
    
End Function

Public Property Get Participants() As List_IParticipant_
    Set Participants = this.Participants
End Property
Public Property Set Participants(ByVal RHS As List_IParticipant_)
    Set this.Participants = RHS
End Property

Public Property Get Worksheet() As ICellFindableExcelWorksheet
    Set Worksheet = this.Worksheet
End Property
Public Property Set Worksheet(ByVal RHS As ICellFindableExcelWorksheet)
    Set this.Worksheet = RHS
End Property

Public Property Get IsReferenceDataUnique() As Boolean
    IsReferenceDataUnique = this.IsReferenceDataUnique
End Property
Public Property Let IsReferenceDataUnique(ByVal RHS As Boolean)
    this.IsReferenceDataUnique = RHS
End Property

Public Property Get ValidationSummary() As List_IComparisonBlockValidMode_
    Set ValidationSummary = this.ValidationSummary
End Property
Public Property Set ValidationSummary(ByVal RHS As List_IComparisonBlockValidMode_)
    Set this.ValidationSummary = RHS
End Property

Public Property Get Self() As IReferenceValuesValidator
    Set Self = Me
End Property

Public Sub Initialize()
    
    If (Participants.Count = vbEmpty) Then _
        Exit Sub
        
    Set ValidationSummary = GetComparisonSummaryBlocks()
    
    IsReferenceDataUnique = GetUniqueStatus()
    
End Sub
    Private Function GetComparisonSummaryBlocks() As List_IComparisonBlockValidMode_
            
        Dim Result As List_IComparisonBlockValidMode_
        Set Result = New List_IComparisonBlockValidMode_
        
        Dim Results As List_IParticipantResult_
        Set Results = GetResults(Participants(1))
        
        Dim BlockCount As Long
        BlockCount = Results.Count
            
        Dim i As Long
        For i = 1 To BlockCount
                
            Dim Model As IComparisonBlockValidationModel
            Set Model = GetBlockModel(i)
                
            Result.Add Model
                
        Next
            
        Set GetComparisonSummaryBlocks = Result
            
    End Function
Private Function GetResults( _
        ByVal Model As IResultsObtainedParticipant) As List_IParticipantResult_
            
    Set GetResults = Model.Results
        
End Function
        Private Function GetBlockModel( _
                ByVal BlockIndex As Long) As IComparisonBlockValidationModel
            
            Dim Model As IResultsObtainedParticipant
            Set Model = Participants(1)
            
            Dim Source As IReferenceTableRowModel
            Set Source = Model.Results(BlockIndex)
            
            Dim AverageValue As Double
            AverageValue = GetReferenceAverageValue(BlockIndex)
            
            Dim MaxDeltaParticipantIndex As Long
            MaxDeltaParticipantIndex = GetMaxDeltaParticipantIndex(BlockIndex, AverageValue)
            
            Dim Result As IComparisonBlockValidationModel
            
            With New ComparisonBlockValidationDTO
                
                .IsReferenceValueUnuque = GetReferenceValueUniqueStatus(BlockIndex, AverageValue)
                
                .BlockIndex = BlockIndex
                .ParticipantIndex = MaxDeltaParticipantIndex
                
                .BlockValueDesctiption = Source.Name & ". " & Source.Subname
                 
                If (MaxDeltaParticipantIndex > vbEmpty = True) Then
                    
                    Dim Organisation As String
                    Organisation = GetOrganization(Participants(MaxDeltaParticipantIndex))
                    
                    .ParticipantName = Organisation
                    
                End If

                Set Result = .Self
                
            End With
            
            Set GetBlockModel = Result
            
        End Function
            Private Function GetOrganization( _
                    ByVal Model As IComparisonDataParticipant) As String
                        
                GetOrganization = Model.Organisation.Value
                            
            End Function
            Private Function GetReferenceAverageValue( _
                    ByVal BlockIndex As Long) As Double
                      
                Dim Calculator As IAverageCalculator
                Set Calculator = New AverageCalculator
                
                Dim i As Long
                For i = 1 To Participants.Count
                                    
                    Dim Results As List_IParticipantResult_
                    Set Results = GetResults(Participants(i))
                    
                    Dim Model As IReferenceTableRowModel
                    Set Model = Results(BlockIndex)
                    
                    Calculator.Append Model.Value
                Next
                
                Dim Result As Double
                Result = Calculator.GetAverage
                            
                GetReferenceAverageValue = Result
                  
            End Function
            Private Function GetReferenceValueUniqueStatus( _
                    ByVal BlockIndex As Long, _
                    ByVal AverageValue As Double) As Boolean
                                         
                Dim Results As List_IParticipantResult_
                Set Results = GetResults(Participants(1))
                                         
                Dim Model As IReferenceTableRowModel
                Set Model = Results(BlockIndex)
                
                Dim Result As Boolean
                Result = Abs(AverageValue - Model.Value) = vbEmpty
                
                GetReferenceValueUniqueStatus = Result
                
            End Function
            Private Function GetMaxDeltaParticipantIndex( _
                    ByVal BlockIndex As Long, _
                    ByVal AverageValue As Double) As Long
                    
                Dim MaxDeltaValue As Double
                Dim MaxDeltaValueParticipantIndex As Long
                
                Dim i As Long
                For i = 1 To Participants.Count
                
                    Dim Results As List_IParticipantResult_
                    Set Results = GetResults(Participants(i))
                
                    Dim Model As IReferenceTableRowModel
                    Set Model = Results(BlockIndex)
                    
                    Dim DeltaValue As Double
                    DeltaValue = Abs(AverageValue - Model.Value)
                    
                    If DeltaValue > MaxDeltaValue Then
                        
                        MaxDeltaValue = DeltaValue
                        MaxDeltaValueParticipantIndex = i
                    
                    End If
                        
                Next
                
                GetMaxDeltaParticipantIndex = MaxDeltaValueParticipantIndex
                    
            End Function
    Private Function GetUniqueStatus() As Boolean
        
        Dim i As Long
        For i = 1 To ValidationSummary.Count
            
            Dim Model As IUniqueableComparisonBlockModel
            Set Model = ValidationSummary(i)
            
            If (Model.IsDataUnique = False) Then _
                Exit Function
                                    
        Next
        
        GetUniqueStatus = True
    
    End Function
'---------------------------------------------
Private Property Get IReferenceValuesValidator_IsReferenceDataUnique() As Boolean
    IReferenceValuesValidator_IsReferenceDataUnique = IsReferenceDataUnique
End Property

Private Sub IReferenceValuesValidator_SuggestCorrection()

    NotifyIncorrectData
    
    MarkNonUniqueCells
    
End Sub
    Private Sub NotifyIncorrectData()

        Dim Notifier As IValidationResultNotifier
        Set Notifier = ValidationResultNotifier.Create(ValidationSummary)
        
        Notifier.SendMessage
        
    End Sub
    Private Sub MarkNonUniqueCells()
        
        Dim i As Long
        For i = 1 To ValidationSummary.Count
            
            Dim Model As IUniqueableComparisonBlockModel
            Set Model = ValidationSummary(i)
            
            If (Model.IsDataUnique = False) Then _
                 MarkCell Model
            
        Next
        
    End Sub
        Private Sub MarkCell( _
                ByVal Model As ISheetSelectableComparisonModel)
            
            If (Worksheet Is Nothing = True) Then _
                Exit Sub
            
            Dim Parser As IExcelCellAddressParser
            Set Parser = ExcelCellAddressParser.Create(Model.BlockIndex, Model.ParticipantIndex)

            Dim TargetCell As IMarkableExcelRange
            Set TargetCell = Worksheet.GetCell(Parser.GetRow, Parser.GetColumn)
            
            TargetCell.Mark
            
        End Sub

Private Property Get IReferenceValuesValidator_ValidationSummary() As List_IComparisonBlockValidMode_
    Set IReferenceValuesValidator_ValidationSummary = ValidationSummary
End Property


