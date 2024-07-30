VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelValuesHandlerFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Transfer Process.Values Handling.Excel"
'@PredeclaredId
Option Explicit
Implements IExcelValuesHandlerFactory

Private Type TExcelValuesHandlerFactory
    
    TemporaryFolderPath As String
    Fso As IFileSystemObject
    
    ReaderFactory As IDocumentReaderFactory
    
End Type

Private this As TExcelValuesHandlerFactory

Public Function Create( _
       ByVal TemporaryFolderPath As String, _
       ByVal Fso As IFileSystemObject, _
       ByVal ReaderFactory As IDocumentReaderFactory) As IExcelValuesHandlerFactory
    
    With New ExcelValuesHandlerFactory
    
        .TemporaryFolderPath = TemporaryFolderPath
        
        Set .Fso = Fso
        Set .ReaderFactory = ReaderFactory
        
        Set Create = .Self
        
    End With
    
End Function

Public Property Get TemporaryFolderPath() As String
    TemporaryFolderPath = this.TemporaryFolderPath
End Property
Public Property Let TemporaryFolderPath(ByVal RHS As String)
    this.TemporaryFolderPath = RHS
End Property

Public Property Get Fso() As IFileSystemObject
    Set Fso = this.Fso
End Property
Public Property Set Fso(ByVal RHS As IFileSystemObject)
    Set this.Fso = RHS
End Property

Public Property Get ReaderFactory() As IDocumentReaderFactory
    Set ReaderFactory = this.ReaderFactory
End Property
Public Property Set ReaderFactory(ByVal RHS As IDocumentReaderFactory)
    Set this.ReaderFactory = RHS
End Property

Public Property Get Self() As IExcelValuesHandlerFactory
    Set Self = Me
End Property

'---------------------------------------------
Private Function IExcelValuesHandlerFactory_Create( _
        ByVal Writer As IDocumentWriter, _
        ByVal Participant As IComparisonDataParticipant) As IValuesHandler
    
    With New ExcelValuesHandler
    
        Set .Writer = Writer
        Set .Participant = Participant
        
        Set .AuxillaryFilesHandlerFactory = GetAuxillaryFilesHandlerFactory(Participant)
    
        Set IExcelValuesHandlerFactory_Create = .Self
        
    End With
    
End Function
    Private Function GetAuxillaryFilesHandlerFactory( _
            ByVal Participant As IComparisonDataParticipant) As IValuesHandlerFactory
            
        Dim TablesFolderPath As String
        TablesFolderPath = Fso.BuildPath(TemporaryFolderPath, Participant.Number.Value)
            
        Dim Tables As List_IFile_
        Set Tables = Fso.GetFolderFiles(TablesFolderPath)
        
        Dim Result As IValuesHandlerFactory
        Set Result = StaticValuesHandlerFactory.Create(Tables, ReaderFactory)
                           
        Set GetAuxillaryFilesHandlerFactory = Result
        
    End Function


