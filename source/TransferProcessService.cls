VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TransferProcessService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Transfer Process"
Option Explicit
Implements ITransferProcessService

Private Type TTransferProcessService

    WriterFactory As IDocumentWriterFactory
    PathBuilder As IDocumentPathBuilderFactory
    
    StaticHandler As IValuesHandlerFactory
    UserInputHandler As IValuesHandlerFactory
    ExcelHandler As IExcelValuesHandlerFactory

    Progress As IProgressServiceDisposer
    InterpretationHandlerFactory As IResultInterpretationHandlerFac
    
    Word As IWordDocumentsService
    Participants As List_IParticipant_
    
End Type

Private this As TTransferProcessService

Public Property Get Self() As ITransferProcessService
    Set Self = Me
End Property

Public Property Get ExcelHandlerFactory() As IExcelValuesHandlerFactory
    Set ExcelHandlerFactory = this.ExcelHandler
End Property
Public Property Set ExcelHandlerFactory(ByVal RHS As IExcelValuesHandlerFactory)
    Set this.ExcelHandler = RHS
End Property

Public Property Get InterpretationHandlerFactory() As IResultInterpretationHandlerFac
    Set InterpretationHandlerFactory = this.InterpretationHandlerFactory
End Property
Public Property Set InterpretationHandlerFactory(ByVal RHS As IResultInterpretationHandlerFac)
    Set this.InterpretationHandlerFactory = RHS
End Property

Public Property Get Participants() As List_IParticipant_
    Set Participants = this.Participants
End Property
Public Property Set Participants(ByVal RHS As List_IParticipant_)
    Set this.Participants = RHS
End Property

Public Property Get PathBuilder() As IDocumentPathBuilderFactory
    Set PathBuilder = this.PathBuilder
End Property
Public Property Set PathBuilder(ByVal RHS As IDocumentPathBuilderFactory)
    Set this.PathBuilder = RHS
End Property

Public Property Get Progress() As IProgressServiceDisposer
    Set Progress = this.Progress
End Property
Public Property Set Progress(ByVal RHS As IProgressServiceDisposer)
    Set this.Progress = RHS
End Property

Public Property Get StaticHandlerFactory() As IValuesHandlerFactory
    Set StaticHandlerFactory = this.StaticHandler
End Property
Public Property Set StaticHandlerFactory(ByVal RHS As IValuesHandlerFactory)
    Set this.StaticHandler = RHS
End Property

Public Property Get UserInputHandler() As IValuesHandlerFactory
    Set UserInputHandler = this.UserInputHandler
End Property
Public Property Set UserInputHandler(ByVal RHS As IValuesHandlerFactory)
    Set this.UserInputHandler = RHS
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

Private Sub Execute( _
        ByVal TemplateDocument As ISavableWordDocument)
    
    If TemplateDocument Is Nothing Then _
       Exit Sub

    Dim Writer As IDocumentWriter
    Set Writer = WriterFactory.Create(TemplateDocument)
    
    Dim Handler As IValuesHandler
    Set Handler = StaticHandlerFactory.Create(Writer)
    
    Handler.ProcessValues
    
    Set Handler = UserInputHandler.Create(Writer)
    Handler.ProcessValues
    
    TemplateDocument.Save
        
    Dim i As Long
    For i = 1 To Participants.Count
        
        Set Handler = ExcelHandlerFactory.Create(Writer, Participants(i))

        Handler.ProcessValues
        
        Dim Results As List_IParticipantResult_
        Set Results = GetResults(Participants(i))
        
        Dim Interpretator As IResultInterpretationHandler
        Set Interpretator = InterpretationHandlerFactory.Create(TemplateDocument, Results)
        
        Interpretator.WriteInterpretation

        CorrectTablesFont TemplateDocument
        UpdateFields TemplateDocument
        
        Dim DocumentName As String
        DocumentName = GetDocumentName(TemplateDocument)
        
        Dim Organisation As String
        Organisation = GetOrganisation(Participants(i))
        
        Dim builder As IDocumentPathBuilder
        Set builder = PathBuilder.Create(DocumentName, Organisation)
        
        Dim sourePath As String
        sourePath = GetDocumentPath(TemplateDocument)
        
        TemplateDocument.SaveAs builder.GetSaveName()
        
        If i = Participants.Count Then _
           Exit Sub

        Set TemplateDocument = Word.GetFrom(sourePath)
        Set Writer = WriterFactory.Create(TemplateDocument)
    
    Next
    
End Sub
    Private Function GetResults( _
            ByVal Model As IResultsObtainedParticipant) As List_IParticipantResult_
            
        Set GetResults = Model.Results
        
    End Function
    Private Function GetDocumentName( _
            ByVal Document As IWordDocument) As String
            
        GetDocumentName = Document.Name
        
    End Function
    Private Function GetOrganisation( _
            ByVal Model As IComparisonDataParticipant) As String
            
        GetOrganisation = Model.Organisation.Value
        
    End Function
        
    Private Function GetDocumentPath( _
            ByVal Document As IWordDocument) As String
            
        GetDocumentPath = Document.Path
        
    End Function
Private Sub CorrectTablesFont( _
            ByVal TemplateDocument As IEditableWordDocument)
            
        Dim Tables As List_IWordTable_
        Set Tables = TemplateDocument.Range.Tables
            
        Dim i As Long
        For i = 1 To Tables.Count
            
            If (Tables(i).IsFilled = True) Then _
                ApplyFontTimesNewRoman Tables(i)
    
        Next
    
    End Sub
        Private Sub ApplyFontTimesNewRoman( _
                ByVal Table As IFormattableWordTable)
            
            Table.ApplyFont "Times New Roman"
            
        End Sub
    Private Function IsTableFilled( _
            ByVal Table As IWordTable) As Boolean
                
        IsTableFilled = Table.IsFilled
    
    End Function
    
    Private Sub UpdateFields( _
            ByVal TemplateDocument As IEditableWordDocument)
            
        Dim Range As IUpdatableWordRange
        Set Range = TemplateDocument.Range
        
        Range.UpdateFields
           
    End Sub
'---------------------------------------------
Private Sub ITransferProcessService_Execute(ByVal TemplateDocument As IWordDocument)
   Execute TemplateDocument
End Sub
Private Sub ITransferProcessService_Dispose()
    
    Word.Dispose
    Progress.Dispose
    
End Sub


