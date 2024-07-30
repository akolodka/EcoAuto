VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ComparisonDecorationServiceFact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Comparison Results Decoration"
'@PredeclaredId
Option Explicit
Implements IComparisonDecorationServiceFac

Private Type TComparisonDecorationServiceFactory

    Fso As IFileSystemObject
    Progress As IProgressService
    
    Word As IWordDocumentsService
    WriterFactory As IDocumentWriterFactory
    
    MeasuredValuesRangeFillerFactory As IValuesRangeFillerFactory
    TemporaryFolderPath As String
    
    Workbooks As IExcelWorkbooksService
    
End Type

Private this As TComparisonDecorationServiceFactory

Public Function Create( _
       ByVal Initial As IComparisonDecoratorInitializat) As IComparisonDecorationServiceFac

    With New ComparisonDecorationServiceFact

        Set .Fso = Initial.Fso
        Set .Progress = Initial.Progress
        
        Set .WriterFactory = Initial.WriterFactory
        Set .Word = Initial.Word
        
        Set .MeasuredValuesRangeFillerFactory = Initial.MeasuredValuesRangeFillerFactory
        .TemporaryFolderPath = Initial.TemporaryFolderPath
        
        Set .Workbooks = Initial.Workbooks
        Set Create = .Self

    End With

End Function

Public Property Get Fso() As IFileSystemObject
    Set Fso = this.Fso
End Property
Public Property Set Fso(ByVal RHS As IFileSystemObject)
    Set this.Fso = RHS
End Property

Public Property Get WriterFactory() As IDocumentWriterFactory
    Set WriterFactory = this.WriterFactory
End Property
Public Property Set WriterFactory(ByVal RHS As IDocumentWriterFactory)
    Set this.WriterFactory = RHS
End Property

Public Property Get Word() As IWordDocumentsService
    Set Word = this.Word
End Property
Public Property Set Word(ByVal RHS As IWordDocumentsService)
    Set this.Word = RHS
End Property

Public Property Get MeasuredValuesRangeFillerFactory() As IValuesRangeFillerFactory
    Set MeasuredValuesRangeFillerFactory = this.MeasuredValuesRangeFillerFactory
End Property
Public Property Set MeasuredValuesRangeFillerFactory(ByVal RHS As IValuesRangeFillerFactory)
    Set this.MeasuredValuesRangeFillerFactory = RHS
End Property

Public Property Get Progress() As IProgressService
    Set Progress = this.Progress
End Property
Public Property Set Progress(ByVal RHS As IProgressService)
    Set this.Progress = RHS
End Property

Public Property Get TemporaryFolderPath() As String
    TemporaryFolderPath = this.TemporaryFolderPath
End Property
Public Property Let TemporaryFolderPath(ByVal RHS As String)
    this.TemporaryFolderPath = RHS
End Property

Public Property Get Workbooks() As IExcelWorkbooksService
    Set Workbooks = this.Workbooks
End Property
Public Property Set Workbooks(ByVal RHS As IExcelWorkbooksService)
    Set this.Workbooks = RHS
End Property

Public Property Get Self() As IComparisonDecorationServiceFac
    Set Self = Me
End Property
'---------------------------------------------
Private Function IComparisonDecorationServiceFac_Create( _
        ByVal Model As IExcelTransferModel) As IComparisonDecorationService

    Dim Filler As IParticipantValuesRangeFiller
    Set Filler = MeasuredValuesRangeFillerFactory.Create( _
                      Model.ControlObjectFactoryNumber, Model.Participants)

    With New ComparisonDecorationService
        
        Set .Fso = Fso
        Set .Participants = Filler.GetParticipantsWithValuesRanges()
        
        .TemporaryFolderPath = TemporaryFolderPath
        Set .Progress = Progress
        
        Set .ParticipantTablesDecoratorFactory = ParticipantTablesDecoratorFacto.Create(Word)
        Set .Workbooks = Workbooks
        
        Set .ComparisonBlockDecoratorFactory = ComparisonBlockDecoratorFactory.Create( _
                Word, WriterFactory, .Participants)

        Set IComparisonDecorationServiceFac_Create = .Self
        
    End With
    
End Function


