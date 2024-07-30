VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DocumentReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("eco_transferToWord.Core.Document IO.Reader")
Option Explicit
Implements IDocumentReader

Private Type TDocumentReader

    FilePath As String
    Reader As ITransferValueReader
    
End Type

Private this As TDocumentReader

Public Property Get Self() As IDocumentReader
    Set Self = Me
End Property

Public Property Get FilePath() As String
    FilePath = this.FilePath
End Property
Public Property Let FilePath(ByVal RHS As String)
    this.FilePath = RHS
End Property

Public Property Get Reader() As ITransferValueReader
    Set Reader = this.Reader
End Property
Public Property Set Reader(ByVal RHS As ITransferValueReader)
    Set this.Reader = RHS
End Property

Private Function ReadContent() As ITransferValue
    
    If (Reader Is Nothing = True) Then _
       Exit Function
    
    Dim Result As ITransferValue
    Set Result = Reader.ReadFrom(FilePath)
    
    Set ReadContent = Result
    
End Function
'---------------------------------------------
Private Function IDocumentReader_ReadContent() As ITransferValue
    Set IDocumentReader_ReadContent = ReadContent
End Function
