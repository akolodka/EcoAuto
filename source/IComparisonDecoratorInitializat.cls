VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IComparisonDecoratorInitializat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("eco_transferToWord.Core.Initialization.Contracts")
'@Interface
Option Explicit

Public Property Get Fso() As IFileSystemObject
End Property

Public Property Get Word() As IWordDocumentsService
End Property

Public Property Get TemporaryFolderPath() As String
End Property

Public Property Get Progress() As IProgressService
End Property

Public Property Get MeasuredValuesRangeFillerFactory() As IValuesRangeFillerFactory
End Property

Public Property Get WriterFactory() As IDocumentWriterFactory
End Property

Public Property Get Workbooks() As IExcelWorkbooksService
End Property
