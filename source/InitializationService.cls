VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InitializationService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("eco_transferToWord.Core.Initialization")
'@PredeclaredId
Option Explicit

Private ProgressPrimary As IProgressServicePrimary
Private WordApplicationFactory As IWordApplicationServiceFactory
Private WorkbooksFactory As IExcelWorkbooksServiceFactory

Public Function Create( _
       Optional ByVal Preset As IInitializationPreset) As IInitializationService
    
    If (Preset Is Nothing = True) Then _
       Set Preset = New InitializationPresetDTO
    
    Set ProgressPrimary = IIf(Preset.ProgressPrimary Is Nothing = True, _
                              ProgressBar.Create(), Preset.ProgressPrimary)
                                    
    Set WordApplicationFactory = IIf(Preset.WordApplicationFactory Is Nothing = True, _
                                     New WordApplicationServiceFactory, Preset.WordApplicationFactory)
    
    Set WorkbooksFactory = IIf(Preset.WorkbooksFactory Is Nothing = True, _
                               New ExcelWorkbooksServiceFactory, Preset.WorkbooksFactory)
    
    With New InitializationDTO

        Set .Configuration = IIf(Preset.Configuration Is Nothing = True, _
                                 ConfigurationService.Create, Preset.Configuration)
                                
        Set .FileSystemObject = IIf(Preset.FileSystemObject Is Nothing = True, _
                                    New FileSystemObjectService, Preset.FileSystemObject)
        '---------------------------------------------
        Set .StaticValueFiles = .FileSystemObject.GetFolderFiles(.Configuration.StaticDataFolderPath)
        Set .DocumentTemplateFiles = .FileSystemObject.GetFolderFiles(.Configuration.TemplatesFolderPath)
        
        Set .SubcontractFiles = .FileSystemObject.GetFolderFiles(.Configuration.SubcontractFolderPath)
        Set .ComparisonAuxillaryFiles = .FileSystemObject.GetFolderFiles(.Configuration.ComparisonAuxillaryFolderPath)

        Set .ControlObjectHandler = ControlObjectHandlerService.Create( _
                                    .FileSystemObject, .Configuration.MeasuredValuesRangesWorkbookPath)
        
        Set .MeasuredValuesRangeFillerFactory = ValuesRangeFillerFactory.Create( _
                                                .FileSystemObject, .Configuration.MeasuredValuesRangesWorkbookPath)
        '---------------------------------------------
        Set .WorkbooksService = WorkbooksFactory.Create(.FileSystemObject)
        Set .SourceSheet = .WorkbooksService.ActiveBook.Worksheets(.Configuration.WorksheetName)
        
        Dim RowsReader As IComparisonDataRowsReader
        Set RowsReader = ComparisonDataRowsReader.Create(.SourceSheet)
        
        Set .ParticipantsHandler = ComparisonDataHandler.Create(RowsReader)
            
        Set .Validator = InitializationValidator.Create(.Self)
        '---------------------------------------------
        Set .CounterFactory = ProgressCounterFactory.Create( _
                              .StaticValueFiles.Count, .ComparisonAuxillaryFiles.Count, .DocumentTemplateFiles.Count)
        
        Set .ProgressPrimary = ProgressPrimary

        Set .PathBuilder = DocumentPathBuilder.Create( _
                           .FileSystemObject, .ProgressPrimary, .Configuration.ResultFolderPath)
                
        Set .WordApp = WordApplicationFactory.Create(.ProgressPrimary)
        
        Set .Closer = DocumentCloser.Create( _
                      .WordApp, .FileSystemObject, .ProgressPrimary)
        
        Set .Cleaner = FolderCleaner.Create( _
                       .FileSystemObject, .ProgressPrimary)
        
        Set .Mover = TemplatesMover.Create( _
                     .FileSystemObject, .ProgressPrimary)
                
        Set .WordDocumentsService = WordDocumentsService.Create( _
                                    .WordApp, .FileSystemObject)
                    
        Set .ResultsPresenter = TransferResultsPresentator.Create( _
                                ExplorerFolderOpener.Create(.FileSystemObject), .Configuration)
        '---------------------------------------------
        Set .TransferValueIOStrategies = New List_ITransferValueIOStrategy_
        
        .TransferValueIOStrategies.Add PlainTextIOStrategy.Create(.FileSystemObject)
        .TransferValueIOStrategies.Add WordRangeIOStrategy.Create(.FileSystemObject, .WordDocumentsService)
        .TransferValueIOStrategies.Add New ChartObjectWriteStrategy
        
        Set .IOFactory = DocumentIOFactory.Create( _
                         .ProgressPrimary, .TransferValueIOStrategies)
                       
        Set .SubcontractHandler = SubcontractValuesHandler.Create( _
                                  .SubcontractFiles, .IOFactory)
        
        Set .StaticHandler = StaticValuesHandlerFactory.Create( _
                             .StaticValueFiles, .IOFactory)
                
        Set .ExcelHandler = ExcelValuesHandlerFactory.Create( _
                            .Configuration.TemporaryFolderPath, .FileSystemObject, .IOFactory)
        '---------------------------------------------
        Set .ResultInterpretationStrategies = New List_IResultInterpretationStra_
        .ResultInterpretationStrategies.Add New BadResultInterpretationStrategy
        .ResultInterpretationStrategies.Add New GoodResultInterpretationStrateg
        
        Set .ResultInterpretationHandlerFactory = ResultInterpretationHandlerFact.Create( _
                                                  .ResultInterpretationStrategies, .ProgressPrimary)
        '---------------------------------------------
        Set .PrepareFactory = TransferPreparationServiceFacto.Create(.Self)
        Set .ProcessFactory = TransferProcessServiceFactory.Create(.Self)
        
        Set .Loader = ResultDocumentTemplateLoader.Create(.Self)
        Set .DecorationServiceFactory = ComparisonDecorationServiceFact.Create(.Self)
        
        Set .Transfer = TransferService.Create(.Self)
        '---------------------------------------------
        Set Create = .Self
        
    End With
    
End Function






