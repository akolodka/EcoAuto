VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProgressCounter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("eco_transferToWord.Core.Initialization.Progress Counter")
Option Explicit
Implements IProgressCounter

Private Type TProgressCounter

    StaticValueFilesCount As Long
    ResultDocumentTemplatesCount As Long
    
    AuxillaryTablesFilesCount As Long
    ParticipantsCount As Long
    
    ComparisonBlockCount As Long
    
End Type

Private this As TProgressCounter

Public Property Get Self() As IProgressCounter
    Set Self = Me
End Property

Public Property Get StaticValueFilesCount() As Long
    StaticValueFilesCount = this.StaticValueFilesCount
End Property
Public Property Let StaticValueFilesCount(ByVal RHS As Long)
    this.StaticValueFilesCount = RHS
End Property

Public Property Get ResultDocumentTemplatesCount() As Long
    ResultDocumentTemplatesCount = this.ResultDocumentTemplatesCount
End Property
Public Property Let ResultDocumentTemplatesCount(ByVal RHS As Long)
    this.ResultDocumentTemplatesCount = RHS
End Property

Public Property Get AuxillaryTablesFilesCount() As Long
    AuxillaryTablesFilesCount = this.AuxillaryTablesFilesCount
End Property
Public Property Let AuxillaryTablesFilesCount(ByVal RHS As Long)
    this.AuxillaryTablesFilesCount = RHS
End Property

Public Property Get ParticipantsCount() As Long
    ParticipantsCount = this.ParticipantsCount
End Property
Public Property Let ParticipantsCount(ByVal RHS As Long)
    this.ParticipantsCount = RHS
End Property

Public Property Get ComparisonBlockCount() As Long
    ComparisonBlockCount = this.ComparisonBlockCount
End Property
Public Property Let ComparisonBlockCount(ByVal RHS As Long)
    this.ComparisonBlockCount = RHS
End Property

Public Function GetProgressTargetCount() As Long
            
    Dim PreparationCount As Long
    PreparationCount = GetPreparationCount
    
    Dim DocumentFillCount As Long
    DocumentFillCount = GetDocumentFillCount
    
    Dim ComparisonDecorationCount As Long
    ComparisonDecorationCount = GetComparisonDecorationCount
    
    '1 + Подключение к word
    '2 + Очистка памяти
    Const WordApplicationCount As Long = 2
    
    Dim Result As Long
    Result = PreparationCount _
             + ComparisonDecorationCount _
             + DocumentFillCount _
             + WordApplicationCount
    
    GetProgressTargetCount = Result

End Function
    Private Function GetPreparationCount() As Long
                    
        Const PreparationUnloadCount As Long = 2
        Const PreparationCleanCount As Long = 2
                    
        '1 - перенос папки основных шаблонов
        'ParticipantsCount - перенос вспомогательных шаблонов
        Dim MoveTemplatesCount As Long
        MoveTemplatesCount = 1 + ParticipantsCount
                    
        Dim MeasuredValuesTableCount As Long
        MeasuredValuesTableCount = ParticipantsCount
                    
        Dim ReferenceValuesTableCount As Long
        ReferenceValuesTableCount = ParticipantsCount
        
        Dim ResultEstimationTableCount As Long
        ResultEstimationTableCount = ParticipantsCount
                    
        Dim Result As Long
        Result = PreparationUnloadCount + _
                 PreparationCleanCount + _
                 MoveTemplatesCount + _
                 MeasuredValuesTableCount + _
                 ReferenceValuesTableCount + _
                 ResultEstimationTableCount
                    
        GetPreparationCount = Result
                    
    End Function
    
    Private Function GetDocumentFillCount() As Long
        
        '1 + субподряд
        '2 + ответственный
        Const UserInputCount As Long = 2
    
        '1 + participant.name
        '2 + participant.number
        '3 + participant.Type
        '4 + Сохранение результатов участника
        '5 + Интерпретация результатов
        
        Const ExcelStageCount As Long = 5
        
        Dim excelValuesFillCount As Long
        excelValuesFillCount = (ExcelStageCount + AuxillaryTablesFilesCount) * ParticipantsCount
        
        Dim Result As Long
        Result = (StaticValueFilesCount + _
                  UserInputCount + _
                  excelValuesFillCount) * ResultDocumentTemplatesCount
                 
        GetDocumentFillCount = Result
    
    End Function
    Private Function GetComparisonDecorationCount() As Long
        
        '1 + таблицы сличений
        '2 + диаграммы сличений
        Const StagesCount As Long = 2
        
        Dim Result As Long
        Result = ParticipantsCount * ComparisonBlockCount * StagesCount
        
        GetComparisonDecorationCount = Result
        
    End Function
'---------------------------------------------
Private Function IProgressCounter_GetProgressTargetCount() As Long
    IProgressCounter_GetProgressTargetCount = GetProgressTargetCount
End Function

