VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITransferProcessInitialization"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("eco_transferToWord.Core.Initialization.Contracts")
'@Interface
Option Explicit

Public Property Get WriterFactory() As IDocumentWriterFactory
End Property

Public Property Get StaticValuesHandler() As IValuesHandlerFactory
End Property

Public Property Get ExcelValuesHandler() As IExcelValuesHandlerFactory
End Property

Public Property Get ProgressPrimary() As IProgressServicePrimary
End Property

Public Property Get CounterFactory() As IProgressCounterFactory
End Property

Public Property Get Word() As IWordDocumentsService
End Property

Public Property Get PathBuilder() As IDocumentPathBuilderFactory
End Property

Public Property Get ResultInterpretationHandlerFactory() As IResultInterpretationHandlerFac
End Property

