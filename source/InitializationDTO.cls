VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InitializationDTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("eco_transferToWord.Core.Initialization")
'@PredeclaredId
Option Explicit

Implements IInitializationService
Implements ITransferProcessInitialization
Implements ITransferPreparationInitializat
Implements ITransferDialogInitialization
Implements IValidatorInitialization
Implements ITemplateLoaderInitialization
Implements ITransferServiceInitialization
Implements IComparisonDecoratorInitializat
Implements IReferenceValuesValidatorInitia

Private Type TInitializationService

    Configuration As IConfiguration
    FileSystemObject As IFileSystemObject
    
    StaticValueFiles As List_IFile_
    DocumentTemplateFiles As List_IFile_
    SubcontractFiles As List_IFile_
    ComparisonAuxillaryFiles As List_IFile_

    ParticipantsHandler As IComparisonDataHandler
    
    Validator As IInitializationValidator
    ControlObjectHandler As IControlObjectHandler

    TransferValueIOStrategies As List_ITransferValueIOStrategy_
    SourceSheet As IExcelWorksheet
    
    CounterFactory As IProgressCounterFactory
    ProgressPrimary As IProgressServicePrimary
    
    PrepareFactory As ITransferPreparationServiceFact
    ProcessFactory As ITransferProcessServiceFactory
    
    Transfer As ITransferService
    IOFactory As IDocumentIOFactory
    
    WordApp As IWordApplicationService
    WordDocumentsService As IWordDocumentsService
    
    StaticHandler As IValuesHandlerFactory
    ExcelHandler As IExcelValuesHandlerFactory
    
    SubcontractHandler As ISubcontractHandler
    PathBuilder As IDocumentPathBuilderFactory
    
    Closer As IDocumentCloser
    Cleaner As IFolderCleaner
    
    Mover As ITemplatesMover
    DecorationServiceFactory As IComparisonDecorationServiceFac
    
    Loader As IResultDocumentTemplateLoader
    MeasuredValuesRangeFillerFactory As IValuesRangeFillerFactory
    
    ResultsPresenter As ITransferResultsPresentator
    WorkbooksService As IExcelWorkbooksService
    
    ResultInterpretationStrategies As List_IResultInterpretationStra_
    ResultInterpretationHandlerFactory As IResultInterpretationHandlerFac
    
End Type

Private this As TInitializationService

Public Property Get Self() As IInitializationService
    Set Self = Me
End Property
'---------------------------------------------

Public Property Get Cleaner() As IFolderCleaner
    Set Cleaner = this.Cleaner
End Property
Public Property Set Cleaner(ByVal RHS As IFolderCleaner)
    Set this.Cleaner = RHS
End Property

Public Property Get Closer() As IDocumentCloser
    Set Closer = this.Closer
End Property
Public Property Set Closer(ByVal RHS As IDocumentCloser)
    Set this.Closer = RHS
End Property

Public Property Get ComparisonAuxillaryFiles() As List_IFile_
    Set ComparisonAuxillaryFiles = this.ComparisonAuxillaryFiles
End Property
Public Property Set ComparisonAuxillaryFiles(ByVal RHS As List_IFile_)
    Set this.ComparisonAuxillaryFiles = RHS
End Property

Public Property Get Configuration() As IConfiguration
    Set Configuration = this.Configuration
End Property
Public Property Set Configuration(ByVal RHS As IConfiguration)
    Set this.Configuration = RHS
End Property

Public Property Get ControlObjectHandler() As IControlObjectHandler
    Set ControlObjectHandler = this.ControlObjectHandler
End Property
Public Property Set ControlObjectHandler(ByVal RHS As IControlObjectHandler)
    Set this.ControlObjectHandler = RHS
End Property

Public Property Get CounterFactory() As IProgressCounterFactory
    Set CounterFactory = this.CounterFactory
End Property
Public Property Set CounterFactory(ByVal RHS As IProgressCounterFactory)
    Set this.CounterFactory = RHS
End Property

Public Property Get DecorationServiceFactory() As IComparisonDecorationServiceFac
    Set DecorationServiceFactory = this.DecorationServiceFactory
End Property
Public Property Set DecorationServiceFactory(ByVal RHS As IComparisonDecorationServiceFac)
    Set this.DecorationServiceFactory = RHS
End Property

Public Property Get DocumentTemplateFiles() As List_IFile_
    Set DocumentTemplateFiles = this.DocumentTemplateFiles
End Property
Public Property Set DocumentTemplateFiles(ByVal RHS As List_IFile_)
    Set this.DocumentTemplateFiles = RHS
End Property

Public Property Get ExcelHandler() As IExcelValuesHandlerFactory
    Set ExcelHandler = this.ExcelHandler
End Property
Public Property Set ExcelHandler(ByVal RHS As IExcelValuesHandlerFactory)
    Set this.ExcelHandler = RHS
End Property

Public Property Get FileSystemObject() As IFileSystemObject
    Set FileSystemObject = this.FileSystemObject
End Property
Public Property Set FileSystemObject(ByVal RHS As IFileSystemObject)
    Set this.FileSystemObject = RHS
End Property

Public Property Get IOFactory() As IDocumentIOFactory
    Set IOFactory = this.IOFactory
End Property
Public Property Set IOFactory(ByVal RHS As IDocumentIOFactory)
    Set this.IOFactory = RHS
End Property

Public Property Get Loader() As IResultDocumentTemplateLoader
    Set Loader = this.Loader
End Property
Public Property Set Loader(ByVal RHS As IResultDocumentTemplateLoader)
    Set this.Loader = RHS
End Property

Public Property Get MeasuredValuesRangeFillerFactory() As IValuesRangeFillerFactory
    Set MeasuredValuesRangeFillerFactory = this.MeasuredValuesRangeFillerFactory
End Property
Public Property Set MeasuredValuesRangeFillerFactory(ByVal RHS As IValuesRangeFillerFactory)
    Set this.MeasuredValuesRangeFillerFactory = RHS
End Property

Public Property Get Mover() As ITemplatesMover
    Set Mover = this.Mover
End Property
Public Property Set Mover(ByVal RHS As ITemplatesMover)
    Set this.Mover = RHS
End Property

Public Property Get ParticipantsHandler() As IComparisonDataHandler
    Set ParticipantsHandler = this.ParticipantsHandler
End Property
Public Property Set ParticipantsHandler(ByVal RHS As IComparisonDataHandler)
    Set this.ParticipantsHandler = RHS
End Property

Public Property Get PathBuilder() As IDocumentPathBuilderFactory
    Set PathBuilder = this.PathBuilder
End Property
Public Property Set PathBuilder(ByVal RHS As IDocumentPathBuilderFactory)
    Set this.PathBuilder = RHS
End Property

Public Property Get PrepareFactory() As ITransferPreparationServiceFact
    Set PrepareFactory = this.PrepareFactory
End Property
Public Property Set PrepareFactory(ByVal RHS As ITransferPreparationServiceFact)
    Set this.PrepareFactory = RHS
End Property

Public Property Get ProcessFactory() As ITransferProcessServiceFactory
    Set ProcessFactory = this.ProcessFactory
End Property
Public Property Set ProcessFactory(ByVal RHS As ITransferProcessServiceFactory)
    Set this.ProcessFactory = RHS
End Property

Public Property Get ProgressPrimary() As IProgressServicePrimary
    Set ProgressPrimary = this.ProgressPrimary
End Property
Public Property Set ProgressPrimary(ByVal RHS As IProgressServicePrimary)
    Set this.ProgressPrimary = RHS
End Property

Public Property Get ResultInterpretationHandlerFactory() As IResultInterpretationHandlerFac
    Set ResultInterpretationHandlerFactory = this.ResultInterpretationHandlerFactory
End Property
Public Property Set ResultInterpretationHandlerFactory(ByVal RHS As IResultInterpretationHandlerFac)
    Set this.ResultInterpretationHandlerFactory = RHS
End Property

Public Property Get ResultInterpretationStrategies() As List_IResultInterpretationStra_
    Set ResultInterpretationStrategies = this.ResultInterpretationStrategies
End Property
Public Property Set ResultInterpretationStrategies(ByVal RHS As List_IResultInterpretationStra_)
    Set this.ResultInterpretationStrategies = RHS
End Property

Public Property Get ResultsPresenter() As ITransferResultsPresentator
    Set ResultsPresenter = this.ResultsPresenter
End Property
Public Property Set ResultsPresenter(ByVal RHS As ITransferResultsPresentator)
    Set this.ResultsPresenter = RHS
End Property

Public Property Get SourceSheet() As IExcelWorksheet
    Set SourceSheet = this.SourceSheet
End Property
Public Property Set SourceSheet(ByVal RHS As IExcelWorksheet)
    Set this.SourceSheet = RHS
End Property

Public Property Get StaticHandler() As IValuesHandlerFactory
    Set StaticHandler = this.StaticHandler
End Property
Public Property Set StaticHandler(ByVal RHS As IValuesHandlerFactory)
    Set this.StaticHandler = RHS
End Property

Public Property Get StaticValueFiles() As List_IFile_
    Set StaticValueFiles = this.StaticValueFiles
End Property
Public Property Set StaticValueFiles(ByVal RHS As List_IFile_)
    Set this.StaticValueFiles = RHS
End Property

Public Property Get SubcontractFiles() As List_IFile_
    Set SubcontractFiles = this.SubcontractFiles
End Property
Public Property Set SubcontractFiles(ByVal RHS As List_IFile_)
    Set this.SubcontractFiles = RHS
End Property

Public Property Get SubcontractHandler() As ISubcontractHandler
    Set SubcontractHandler = this.SubcontractHandler
End Property
Public Property Set SubcontractHandler(ByVal RHS As ISubcontractHandler)
    Set this.SubcontractHandler = RHS
End Property

Public Property Get Transfer() As ITransferService
    Set Transfer = this.Transfer
End Property
Public Property Set Transfer(ByVal RHS As ITransferService)
    Set this.Transfer = RHS
End Property

Public Property Get TransferValueIOStrategies() As List_ITransferValueIOStrategy_
    Set TransferValueIOStrategies = this.TransferValueIOStrategies
End Property
Public Property Set TransferValueIOStrategies(ByVal RHS As List_ITransferValueIOStrategy_)
    Set this.TransferValueIOStrategies = RHS
End Property

Public Property Get Validator() As IInitializationValidator
    Set Validator = this.Validator
End Property
Public Property Set Validator(ByVal RHS As IInitializationValidator)
    Set this.Validator = RHS
End Property

Public Property Get WordApp() As IWordApplicationService
    Set WordApp = this.WordApp
End Property
Public Property Set WordApp(ByVal RHS As IWordApplicationService)
    Set this.WordApp = RHS
End Property

Public Property Get WordDocumentsService() As IWordDocumentsService
    Set WordDocumentsService = this.WordDocumentsService
End Property
Public Property Set WordDocumentsService(ByVal RHS As IWordDocumentsService)
    Set this.WordDocumentsService = RHS
End Property

Public Property Get WorkbooksService() As IExcelWorkbooksService
    Set WorkbooksService = this.WorkbooksService
End Property
Public Property Set WorkbooksService(ByVal RHS As IExcelWorkbooksService)
    Set this.WorkbooksService = RHS
End Property
'---------------------------------------------
Private Property Get ITransferProcessInitialization_WriterFactory() As IDocumentWriterFactory
    Set ITransferProcessInitialization_WriterFactory = IOFactory
End Property
Private Property Get ITransferProcessInitialization_StaticValuesHandler() As IValuesHandlerFactory
    Set ITransferProcessInitialization_StaticValuesHandler = StaticHandler
End Property
Private Property Get ITransferProcessInitialization_ExcelValuesHandler() As IExcelValuesHandlerFactory
    Set ITransferProcessInitialization_ExcelValuesHandler = ExcelHandler
End Property
Private Property Get ITransferProcessInitialization_ProgressPrimary() As IProgressServicePrimary
    Set ITransferProcessInitialization_ProgressPrimary = ProgressPrimary
End Property
Private Property Get ITransferProcessInitialization_Word() As IWordDocumentsService
    Set ITransferProcessInitialization_Word = WordDocumentsService
End Property
Private Property Get ITransferProcessInitialization_PathBuilder() As IDocumentPathBuilderFactory
    Set ITransferProcessInitialization_PathBuilder = PathBuilder
End Property
Private Property Get ITransferProcessInitialization_CounterFactory() As IProgressCounterFactory
    Set ITransferProcessInitialization_CounterFactory = CounterFactory
End Property
Private Property Get ITransferProcessInitialization_ResultInterpretationHandlerFactory() As IResultInterpretationHandlerFac
    Set ITransferProcessInitialization_ResultInterpretationHandlerFactory = ResultInterpretationHandlerFactory
End Property
'---------------------------------------------
Private Property Get ITransferPreparationInitializat_Config() As ITransferPreparationConfig
    Set ITransferPreparationInitializat_Config = Configuration
End Property
Private Property Get ITransferPreparationInitializat_Fso() As IFileSystemObject
    Set ITransferPreparationInitializat_Fso = FileSystemObject
End Property
Private Property Get ITransferPreparationInitializat_Closer() As IDocumentCloser
    Set ITransferPreparationInitializat_Closer = Closer
End Property
Private Property Get ITransferPreparationInitializat_Cleaner() As IFolderCleaner
    Set ITransferPreparationInitializat_Cleaner = Cleaner
End Property
Private Property Get ITransferPreparationInitializat_Mover() As ITemplatesMover
    Set ITransferPreparationInitializat_Mover = Mover
End Property
'---------------------------------------------
Private Property Get ITransferDialogInitialization_SubcontractFiles() As List_IFile_
    Set ITransferDialogInitialization_SubcontractFiles = SubcontractFiles
End Property
Private Property Get ITransferDialogInitialization_Participants() As List_IParticipant_
    Set ITransferDialogInitialization_Participants = ParticipantsHandler.GetParticipants
End Property
Private Property Get ITransferDialogInitialization_SubcontractHandler() As ISubcontractHandler
    Set ITransferDialogInitialization_SubcontractHandler = SubcontractHandler
End Property
Private Property Get ITransferDialogInitialization_ValidationStatus() As IValidationStatus
    Set ITransferDialogInitialization_ValidationStatus = Validator.GetValidationStatus
End Property
Private Property Get ITransferDialogInitialization_TransferService() As ITransferService
    Set ITransferDialogInitialization_TransferService = Transfer
End Property
Private Property Get ITransferDialogInitialization_FileSystemObject() As IFileSystemObject
    Set ITransferDialogInitialization_FileSystemObject = FileSystemObject
End Property
Private Property Get ITransferDialogInitialization_ControlObjectHandler() As IControlObjectHandler
    Set ITransferDialogInitialization_ControlObjectHandler = ControlObjectHandler
End Property
'---------------------------------------------
Private Property Get IValidatorInitialization_TemplateFilesCount() As Long
    IValidatorInitialization_TemplateFilesCount = DocumentTemplateFiles.Count
End Property
Private Property Get IValidatorInitialization_StaticValueFilesCount() As Long
    IValidatorInitialization_StaticValueFilesCount = StaticValueFiles.Count
End Property
Private Property Get IValidatorInitialization_SubctontractFilesCount() As Long
    IValidatorInitialization_SubctontractFilesCount = SubcontractFiles.Count
End Property
Private Property Get IValidatorInitialization_ComparisonAuxillaryFilesCount() As Long
    IValidatorInitialization_ComparisonAuxillaryFilesCount = ComparisonAuxillaryFiles.Count
End Property
Private Property Get IValidatorInitialization_ParticipantsCount() As Long
    IValidatorInitialization_ParticipantsCount = ParticipantsHandler.GetParticipants().Count
End Property
'---------------------------------------------
Private Property Get ITemplateLoaderInitialization_Fso() As IFileSystemObject
    Set ITemplateLoaderInitialization_Fso = FileSystemObject
End Property
Private Property Get ITemplateLoaderInitialization_Word() As IWordDocumentsService
    Set ITemplateLoaderInitialization_Word = WordDocumentsService
End Property
Private Property Get ITemplateLoaderInitialization_TemporaryFolderPath() As String
    ITemplateLoaderInitialization_TemporaryFolderPath = Configuration.TemporaryFolderPath
End Property
'---------------------------------------------
Private Property Get ITransferServiceInitialization_PreparerFactory() As ITransferPreparationServiceFact
    Set ITransferServiceInitialization_PreparerFactory = PrepareFactory
End Property
Private Property Get ITransferServiceInitialization_ProcessFactory() As ITransferProcessServiceFactory
    Set ITransferServiceInitialization_ProcessFactory = ProcessFactory
End Property
Private Property Get ITransferServiceInitialization_Loader() As IResultDocumentTemplateLoader
    Set ITransferServiceInitialization_Loader = Loader
End Property
Private Property Get ITransferServiceInitialization_WordApp() As IWordApplicationQuitService
    Set ITransferServiceInitialization_WordApp = WordApp
End Property
Private Property Get ITransferServiceInitialization_ResultPresenter() As ITransferResultsPresentator
    Set ITransferServiceInitialization_ResultPresenter = ResultsPresenter
End Property
Private Property Get ITransferServiceInitialization_DecoratorFactory() As IComparisonDecorationServiceFac
    Set ITransferServiceInitialization_DecoratorFactory = DecorationServiceFactory
End Property
Private Property Get ITransferServiceInitialization_Progress() As IProgressServiceDisposer
    Set ITransferServiceInitialization_Progress = ProgressPrimary
End Property
'---------------------------------------------
Private Property Get IComparisonDecoratorInitializat_Fso() As IFileSystemObject
    Set IComparisonDecoratorInitializat_Fso = FileSystemObject
End Property
Private Property Get IComparisonDecoratorInitializat_Word() As IWordDocumentsService
    Set IComparisonDecoratorInitializat_Word = WordDocumentsService
End Property
Private Property Get IComparisonDecoratorInitializat_TemporaryFolderPath() As String
    IComparisonDecoratorInitializat_TemporaryFolderPath = Configuration.TemporaryFolderPath
End Property
Private Property Get IComparisonDecoratorInitializat_Progress() As IProgressService
    Set IComparisonDecoratorInitializat_Progress = ProgressPrimary
End Property
Private Property Get IComparisonDecoratorInitializat_MeasuredValuesRangeFillerFactory() As IValuesRangeFillerFactory
    Set IComparisonDecoratorInitializat_MeasuredValuesRangeFillerFactory = MeasuredValuesRangeFillerFactory
End Property
Private Property Get IComparisonDecoratorInitializat_WriterFactory() As IDocumentWriterFactory
    Set IComparisonDecoratorInitializat_WriterFactory = IOFactory
End Property
Private Property Get IComparisonDecoratorInitializat_Workbooks() As IExcelWorkbooksService
    Set IComparisonDecoratorInitializat_Workbooks = WorkbooksService
End Property
'---------------------------------------------
Private Property Get IReferenceValuesValidatorInitia_Participants() As List_IParticipant_
    Set IReferenceValuesValidatorInitia_Participants = ParticipantsHandler.GetParticipants()
End Property
Private Property Get IReferenceValuesValidatorInitia_SourceWorksheet() As IExcelWorksheet
    Set IReferenceValuesValidatorInitia_SourceWorksheet = SourceSheet
End Property
