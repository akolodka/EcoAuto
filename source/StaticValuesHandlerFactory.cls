VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StaticValuesHandlerFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Transfer Process.Values Handling.Static"
'@PredeclaredId
Option Explicit
Implements IValuesHandlerFactory

Private Type TStaticValuesHandlerFactory
    
    StaticFiles As List_IFile_
    ReaderFactory As IDocumentReaderFactory
    
End Type

Private this As TStaticValuesHandlerFactory

Public Function Create( _
       ByVal StaticFiles As List_IFile_, _
       ByVal ReaderFactory As IDocumentReaderFactory) As IValuesHandlerFactory

    With New StaticValuesHandlerFactory
        
        Set .StaticFiles = StaticFiles
        Set .ReaderFactory = ReaderFactory
        
        Set Create = .Self
        
    End With
    
End Function

Public Property Get StaticFiles() As List_IFile_
    Set StaticFiles = this.StaticFiles
End Property
Public Property Set StaticFiles(ByVal RHS As List_IFile_)
    Set this.StaticFiles = RHS
End Property

Public Property Get ReaderFactory() As IDocumentReaderFactory
    Set ReaderFactory = this.ReaderFactory
End Property
Public Property Set ReaderFactory(ByVal RHS As IDocumentReaderFactory)
    Set this.ReaderFactory = RHS
End Property

Public Property Get Self() As IValuesHandlerFactory
    Set Self = Me
End Property
'---------------------------------------------
Private Function IValuesHandlerFactory_Create( _
        ByVal Writer As IDocumentWriter) As IValuesHandler
    
    With New StaticValuesHandler
    
        Set .KeyValuePairs = GetPairList
        Set .Writer = Writer
        
        Set IValuesHandlerFactory_Create = .Self
    
    End With
    
End Function
    Private Function GetPairList() As List_IKeyValuePair_
            
        Dim Result As List_IKeyValuePair_
        Set Result = New List_IKeyValuePair_
            
        Dim i As Long
        For i = 1 To StaticFiles.Count
            
            Dim Reader As IDocumentReader
            Set Reader = ReaderFactory.Create(StaticFiles(i).Path)
                
            Dim Value As ITransferValue
            Set Value = Reader.ReadContent()
                
            Dim Pair As IKeyValuePair
            Set Pair = KeyValuePair.Create( _
                       StaticFiles(i).NameWithoutExtension, Value)
                
            Result.Add Pair
            
        Next
            
        Set GetPairList = Result
            
    End Function
