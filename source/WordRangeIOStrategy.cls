VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WordRangeIOStrategy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Document IO.IO Strategies"
'@PredeclaredId
Option Explicit

Implements ITransferValueIOStrategy
Implements ITransferValueReader
Implements ITransferValueWriter

Const DefaultIOExtensions As String = "doc docx"

Private Type TWordRangeIOStrategy

    Fso As IFileSystemObject
    Word As IWordDocumentsService
    
End Type

Private this As TWordRangeIOStrategy

Public Function Create( _
       ByVal Fso As IFileSystemObject, _
       ByVal Word As IWordDocumentsService) As ITransferValueIOStrategy
    
    With New WordRangeIOStrategy
    
        Set .Fso = Fso
        Set .Word = Word
        
        Set Create = .Self
    
    End With
    
End Function

Public Property Get Fso() As IFileSystemObject
    Set Fso = this.Fso
End Property
Public Property Set Fso(ByVal RHS As IFileSystemObject)
    Set this.Fso = RHS
End Property

Public Property Get Word() As IWordDocumentsService
    Set Word = this.Word
End Property
Public Property Set Word(ByVal RHS As IWordDocumentsService)
    Set this.Word = RHS
End Property

Public Property Get Self() As ITransferValueIOStrategy
    Set Self = Me
End Property

Public Function IsMatchesForRead( _
       ByVal FilePath As String) As Boolean
    
    Dim extensions() As String
    extensions = Split(DefaultIOExtensions, " ")
    
    Dim fileExtension As String
    fileExtension = Fso.GetExtensionName(FilePath)

    Dim i As Long
    For i = LBound(extensions) To UBound(extensions)
        
        If extensions(i) = fileExtension Then
        
            IsMatchesForRead = True
            Exit Function
        
        End If
        
    Next
        
End Function

Public Function ReadFrom( _
       ByVal FilePath As String) As ITransferValue
    
    If (IsMatchesForRead(FilePath) = False) Then _
       Exit Function
    
    Dim Document As IEditableWordDocument
    Set Document = Word.GetFrom(FilePath)
    
    Dim Result As ITransferValue
    Set Result = TransferValue.Create("WordDocument value not readed")
    
    If (Document Is Nothing = False) Then _
       Set Result = TransferValue.Create(Document.Range)
    
    Set ReadFrom = Result
    
End Function

Public Function IsMatchesForWrite( _
       ByVal Value As ITransferValue) As Boolean
    
    If (Value Is Nothing = True) Then _
       Exit Function
    
    If (IsObject(Value.Value) = False) Then _
       Exit Function
        
    If (Value.Value Is Nothing = True) Then _
       Exit Function
        
    IsMatchesForWrite = TypeOf Value.Value Is IWordRange
    
End Function

Public Sub InsertIntoDocument( _
       ByVal Value As ITransferValue, _
       ByVal Destination As IEditableWordRange)
    
    If (IsMatchesForWrite(Value) = False) Then _
       Exit Sub
    
    If (Destination Is Nothing = True) Then _
       Exit Sub
    
    Dim Source As IEditableWordRange
    Set Source = Value.Value
        
    If Source Is Nothing Then _
       Exit Sub
    
    Source.Copy
    Destination.PasteAndFormat
    
End Sub

'---------------------------------------------
Private Function ITransferValueIOStrategy_IsMatchesForRead( _
        ByVal FilePath As String) As Boolean
    
    ITransferValueIOStrategy_IsMatchesForRead = IsMatchesForRead(FilePath)
    
End Function

Private Function ITransferValueIOStrategy_IsMatchesForWrite( _
        ByVal Value As ITransferValue) As Boolean
    
    ITransferValueIOStrategy_IsMatchesForWrite = IsMatchesForWrite(Value)
End Function

'---------------------------------------------

Private Function ITransferValueReader_ReadFrom( _
        ByVal FilePath As String) As ITransferValue
    
    Set ITransferValueReader_ReadFrom = ReadFrom(FilePath)

End Function

'---------------------------------------------
Private Sub ITransferValueWriter_InsertIntoDocument( _
        ByVal Value As ITransferValue, _
        ByVal Destination As IWordRange)
    
    InsertIntoDocument Value, Destination
    
End Sub


