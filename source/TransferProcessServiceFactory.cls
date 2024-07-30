VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TransferProcessServiceFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("eco_transferToWord.Core.Transfer Process")
'@PredeclaredId
Option Explicit
Implements ITransferProcessServiceFactory

Private Type TTransferProcessFactory
    
    ProgressPrimary As IProgressServicePrimary
    CounterFactory As IProgressCounterFactory

    WriterFactory As IDocumentWriterFactory
    
    StaticHandler As IValuesHandlerFactory
    ExcelHandler As IExcelValuesHandlerFactory

    Word As IWordDocumentsService
    PathBuilder As IDocumentPathBuilderFactory
    
    InterpretationHandlerFactory As IResultInterpretationHandlerFac
    
End Type

Private this As TTransferProcessFactory

Public Function Create( _
       ByVal Initial As ITransferProcessInitialization) As ITransferProcessServiceFactory
        
    With New TransferProcessServiceFactory
        
        Set .ProgressPrimary = Initial.ProgressPrimary
        Set .CounterFactory = Initial.CounterFactory
        
        Set .WriterFactory = Initial.WriterFactory
        
        Set .StaticHandler = Initial.StaticValuesHandler
        Set .ExcelHandler = Initial.ExcelValuesHandler
        
        Set .Word = Initial.Word
        Set .PathBuilder = Initial.PathBuilder
        
        Set .InterpretationHandlerFactory = Initial.ResultInterpretationHandlerFactory
        
        Set Create = .Self
        
    End With
    
End Function

Public Property Get CounterFactory() As IProgressCounterFactory
    Set CounterFactory = this.CounterFactory
End Property
Public Property Set CounterFactory(ByVal RHS As IProgressCounterFactory)
    Set this.CounterFactory = RHS
End Property

Public Property Get ExcelHandler() As IExcelValuesHandlerFactory
    Set ExcelHandler = this.ExcelHandler
End Property
Public Property Set ExcelHandler(ByVal RHS As IExcelValuesHandlerFactory)
    Set this.ExcelHandler = RHS
End Property

Public Property Get InterpretationHandlerFactory() As IResultInterpretationHandlerFac
    Set InterpretationHandlerFactory = this.InterpretationHandlerFactory
End Property
Public Property Set InterpretationHandlerFactory(ByVal RHS As IResultInterpretationHandlerFac)
    Set this.InterpretationHandlerFactory = RHS
End Property

Public Property Get PathBuilder() As IDocumentPathBuilderFactory
    Set PathBuilder = this.PathBuilder
End Property
Public Property Set PathBuilder(ByVal RHS As IDocumentPathBuilderFactory)
    Set this.PathBuilder = RHS
End Property

Public Property Get ProgressPrimary() As IProgressServicePrimary
    Set ProgressPrimary = this.ProgressPrimary
End Property
Public Property Set ProgressPrimary(ByVal RHS As IProgressServicePrimary)
    Set this.ProgressPrimary = RHS
End Property

Public Property Get StaticHandler() As IValuesHandlerFactory
    Set StaticHandler = this.StaticHandler
End Property
Public Property Set StaticHandler(ByVal RHS As IValuesHandlerFactory)
    Set this.StaticHandler = RHS
End Property

Public Property Get Word() As IWordDocumentsService
    Set Word = this.Word
End Property
Public Property Set Word(ByVal RHS As IWordDocumentsService)
    Set this.Word = RHS
End Property

Public Property Get WriterFactory() As IDocumentWriterFactory
    Set WriterFactory = this.WriterFactory
End Property
Public Property Set WriterFactory(ByVal RHS As IDocumentWriterFactory)
    Set this.WriterFactory = RHS
End Property

Public Property Get Self() As ITransferProcessServiceFactory
    Set Self = Me
End Property
'---------------------------------------------
Private Function ITransferProcessServiceFactory_Create( _
        ByVal Model As ITransferMenuModel) As ITransferProcessService
    
    Dim Participants As List_IParticipant_
    Set Participants = GetParticipants(Model)
    
    Dim Results As List_IParticipantResult_
    Set Results = GetResults(Participants(1))
    
    Dim Counter As IProgressCounter
    Set Counter = CounterFactory.Create(Participants.Count, Results.Count)
    
    Dim TotalCount As Long
    TotalCount = Counter.GetProgressTargetCount()
    
    With New TransferProcessService
    
        Set .WriterFactory = WriterFactory
        Set .InterpretationHandlerFactory = InterpretationHandlerFactory
        
        Set .StaticHandlerFactory = StaticHandler
        Set .ExcelHandlerFactory = ExcelHandler
        
        Set .Word = Word
        Set .PathBuilder = PathBuilder
    
        Set .UserInputHandler = UserInputValuesHandlerFactory.Create(Model)
        Set .Participants = Participants
        
        Set .Progress = ProgressPrimary.Initialize(TotalCount)
    
        Set ITransferProcessServiceFactory_Create = .Self
        
    End With

End Function
    Private Function GetParticipants( _
            ByVal Model As IExcelTransferModel) As List_IParticipant_
            
        Set GetParticipants = Model.Participants
            
    End Function
    Private Function GetResults( _
            ByVal Model As IResultsObtainedParticipant) As List_IParticipantResult_
                
        Set GetResults = Model.Results
                
    End Function



