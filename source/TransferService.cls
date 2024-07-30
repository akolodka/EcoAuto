VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TransferService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("eco_transferToWord.Core")
'@PredeclaredId
Option Explicit
Implements ITransferService

Private Type TTransferService

    PreparerFactory As ITransferPreparationServiceFact
    TransferFactory As ITransferProcessServiceFactory
    
    Loader As IResultDocumentTemplateLoader
    WordApp As IWordApplicationQuitService
    
    ResultsPresenter As ITransferResultsPresentator
    DecorationFactory As IComparisonDecorationServiceFac
    
    Progress As IProgressServiceDisposer
    
End Type

Private this As TTransferService

Public Function Create( _
       ByVal Initial As ITransferServiceInitialization) As ITransferService
    
    With New TransferService
        
        Set .PreparerFactory = Initial.PreparerFactory
        Set .TransferFactory = Initial.ProcessFactory
        
        Set .Loader = Initial.Loader
        Set .WordApp = Initial.WordApp
        
        Set .ResultsPresenter = Initial.ResultPresenter
        Set .DecorationFactory = Initial.DecoratorFactory
        
        Set .Progress = Initial.Progress
        Set Create = .Self
        
    End With

End Function

Public Property Get PreparerFactory() As ITransferPreparationServiceFact
    Set PreparerFactory = this.PreparerFactory
End Property
Public Property Set PreparerFactory(ByVal RHS As ITransferPreparationServiceFact)
    Set this.PreparerFactory = RHS
End Property

Public Property Get Progress() As IProgressServiceDisposer
    Set Progress = this.Progress
End Property
Public Property Set Progress(ByVal RHS As IProgressServiceDisposer)
    Set this.Progress = RHS
End Property

Public Property Get TransferFactory() As ITransferProcessServiceFactory
    Set TransferFactory = this.TransferFactory
End Property
Public Property Set TransferFactory(ByVal RHS As ITransferProcessServiceFactory)
    Set this.TransferFactory = RHS
End Property

Public Property Get DecorationFactory() As IComparisonDecorationServiceFac
    Set DecorationFactory = this.DecorationFactory
End Property
Public Property Set DecorationFactory(ByVal RHS As IComparisonDecorationServiceFac)
    Set this.DecorationFactory = RHS
End Property

Public Property Get Loader() As IResultDocumentTemplateLoader
    Set Loader = this.Loader
End Property
Public Property Set Loader(ByVal RHS As IResultDocumentTemplateLoader)
    Set this.Loader = RHS
End Property

Public Property Get ResultsPresenter() As ITransferResultsPresentator
    Set ResultsPresenter = this.ResultsPresenter
End Property
Public Property Set ResultsPresenter(ByVal RHS As ITransferResultsPresentator)
    Set this.ResultsPresenter = RHS
End Property

Public Property Get WordApp() As IWordApplicationQuitService
    Set WordApp = this.WordApp
End Property
Public Property Set WordApp(ByVal RHS As IWordApplicationQuitService)
    Set this.WordApp = RHS
End Property

Public Property Get Self() As ITransferService
    Set Self = Me
End Property
'---------------------------------------------
Private Sub ITransferService_Execute( _
        ByVal Model As ITransferMenuModel)

    On Error GoTo CleanFail
    
    'Идёт первым -- внутри инициализируется прогрессбар
    Dim Transfer As ITransferProcessService
    Set Transfer = TransferFactory.Create(Model)
    
    Dim Preparer As ITransferPreparationService
    Set Preparer = PreparerFactory.Create(Model)
    
    Preparer.PrepareTemplates
    
    Dim Decorator As IComparisonDecorationService
    Set Decorator = DecorationFactory.Create(Model)
    
    Decorator.DecorateComparisonResults
    
    Dim Templates As List_IWordDocument_
    Set Templates = Loader.GetTemplates()

    Dim i As Long
    For i = 1 To Templates.Count

        Transfer.Execute Templates(i)

    Next
    
    Transfer.Dispose
    
    ResultsPresenter.OpenResultsFolder
    
    Exit Sub
        
CleanFail:
    
    MsgBox "Возникла неустранимая ошибка:" & vbNewLine & Err.Number & Err.Description, vbCritical
    
    WordApp.EmergencyQuit
    Progress.Dispose

End Sub


