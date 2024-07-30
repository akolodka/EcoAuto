VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITransferServiceInitialization"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("eco_transferToWord.Core.Initialization.Contracts")
'@Interface
Option Explicit

Public Property Get PreparerFactory() As ITransferPreparationServiceFact
End Property

Public Property Get ProcessFactory() As ITransferProcessServiceFactory
End Property

Public Property Get Loader() As IResultDocumentTemplateLoader
End Property

Public Property Get WordApp() As IWordApplicationQuitService
End Property

Public Property Get ResultPresenter() As ITransferResultsPresentator
End Property

Public Property Get DecoratorFactory() As IComparisonDecorationServiceFac
End Property

Public Property Get Progress() As IProgressServiceDisposer
End Property
