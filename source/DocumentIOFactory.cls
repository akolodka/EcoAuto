VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DocumentIOFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Document IO"
'@PredeclaredId
Option Explicit

Implements IDocumentIOFactory
Implements IDocumentReaderFactory
Implements IDocumentWriterFactory

Private Type TDocumentIOFactory

    Progress As IProgressService
    StrategiesIO As List_ITransferValueIOStrategy_
    
End Type

Private this As TDocumentIOFactory

Public Function Create( _
       ByVal Progress As IProgressService, _
       ByVal Strategies As List_ITransferValueIOStrategy_) As IDocumentIOFactory
    
    With New DocumentIOFactory
        
        Set .Progress = Progress
        Set .StrategiesIO = Strategies
        
        Set Create = .Self
        
    End With
    
End Function

Public Property Get Progress() As IProgressService
    Set Progress = this.Progress
End Property
Public Property Set Progress(ByVal RHS As IProgressService)
    Set this.Progress = RHS
End Property

Public Property Get StrategiesIO() As List_ITransferValueIOStrategy_
    Set StrategiesIO = this.StrategiesIO
End Property
Public Property Set StrategiesIO(ByVal RHS As List_ITransferValueIOStrategy_)
    Set this.StrategiesIO = RHS
End Property

Public Property Get Self() As IDocumentIOFactory
    Set Self = Me
End Property
'---------------------------------------------
Private Function IDocumentReaderFactory_Create( _
        ByVal FilePath As String) As IDocumentReader
    
    With New DocumentReader
        
        .FilePath = FilePath
        Set .Reader = GetReadStrategy(FilePath)
        
        Set IDocumentReaderFactory_Create = .Self
        
    End With
    
End Function
    Private Function GetReadStrategy( _
            ByVal FilePath As String) As ITransferValueReader
            
        Dim i As Long
        For i = 1 To StrategiesIO.Count
            
            If (StrategiesIO(i).IsMatchesForRead(FilePath) = True) Then
                
                Set GetReadStrategy = StrategiesIO(i)
                Exit Function
                    
            End If
                
        Next
            
    End Function
'---------------------------------------------
Private Function IDocumentWriterFactory_Create( _
        ByVal Document As IWordDocument) As IDocumentWriter
    
    With New DocumentWriter
        
        Set .Progress = Progress
        Set .StrategiesIO = StrategiesIO
        
        Set .Finder = DocumentRangeFinder.Create(Document)
        
        If (Document Is Nothing = False) Then _
            .DocumentName = Document.Name
    
        Set IDocumentWriterFactory_Create = .Self
        
    End With

End Function


