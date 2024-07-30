VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WordDocumentsService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Word"
'@PredeclaredId
Option Explicit
Implements IWordDocumentsService

Private Type TWordDocumentsService

    ApplicationWord As IWordApplicationService
    Fso As IFileSystemObject
    
    DocumentsWord As List_IWordDocument_
    IsTempDocumentObtained As Boolean
    
    TemporaryDocument As IEditableWordDocument
    
End Type

Private this As TWordDocumentsService

Public Function Create( _
       ByVal ApplicationWord As IWordApplicationService, _
       ByVal FileSystemObject As IFileSystemObject) As IWordDocumentsService
    
    With New WordDocumentsService

        Set .ApplicationWord = ApplicationWord
        Set .Fso = FileSystemObject
        
        Set Create = .Self
        
    End With

End Function

Public Property Get ApplicationWord() As IWordApplicationService
    Set ApplicationWord = this.ApplicationWord
End Property
Public Property Set ApplicationWord(ByVal RHS As IWordApplicationService)
    Set this.ApplicationWord = RHS
End Property

Public Property Get DocumentsWord() As List_IWordDocument_
    Set DocumentsWord = this.DocumentsWord
End Property
Public Property Set DocumentsWord(ByVal RHS As List_IWordDocument_)
    Set this.DocumentsWord = RHS
End Property

Public Property Get Fso() As IFileSystemObject
    Set Fso = this.Fso
End Property
Public Property Set Fso(ByVal RHS As IFileSystemObject)
    Set this.Fso = RHS
End Property

Public Property Get Self() As IWordDocumentsService

    Set DocumentsWord = New List_IWordDocument_
    Set Self = Me
    
End Property

Public Property Get TemporaryDocument() As IEditableWordDocument
    Set TemporaryDocument = this.TemporaryDocument
End Property
Public Property Set TemporaryDocument(ByVal RHS As IEditableWordDocument)
    Set this.TemporaryDocument = RHS
End Property

Private Function GetTemporaryDocument() As IWordDocument
    
    If (IsTempDocumentObtained = False) Then _
        Set TemporaryDocument = GetNewDocument
    
    TemporaryDocument.Range = vbNullString
    
    Set GetTemporaryDocument = TemporaryDocument
    
End Function
    Private Function GetNewDocument() As IWordDocument
        
        Dim Result As IWordDocument
        Set Result = ApplicationWord.DocumentsAdd(Visible:=False)
        
        DocumentsWord.Add Result
        Set GetNewDocument = Result
        
    End Function
    

Public Property Get IsTempDocumentObtained() As Boolean
    IsTempDocumentObtained = this.IsTempDocumentObtained
End Property
Public Property Let IsTempDocumentObtained(ByVal RHS As Boolean)
    this.IsTempDocumentObtained = RHS
End Property

'---------------------------------------------
Private Function IWordDocumentsService_GetTemporaryDocument() As IWordDocument
    Set IWordDocumentsService_GetTemporaryDocument = GetTemporaryDocument
End Function

Private Function IWordDocumentsService_GetFrom( _
        ByVal Path As String) As IWordDocument
    
    If (DocumentsWord Is Nothing = True) Then _
       Exit Function
    
    If (IsDocumentReadable(Path) = False) Then _
       Exit Function
    
    Dim Document As IWordDocument
    If (DocumentsWord.Count = vbEmpty) Then
        
        Set Document = LoadFrom(Path)
        DocumentsWord.Add Document

        Set IWordDocumentsService_GetFrom = Document
        Exit Function
    
    End If
    
    Dim i As Long
    For i = 1 To DocumentsWord.Count
        
        If (DocumentsWord(i) Is Nothing = True) Then _
           Exit Function
        
        If (Path = DocumentsWord(i).ParentFolderPath) Then
            
            Set IWordDocumentsService_GetFrom = DocumentsWord(i)
            Exit Function
            
        End If
        
    Next
    
    Set Document = LoadFrom(Path)
    DocumentsWord.Add Document

    Set IWordDocumentsService_GetFrom = Document
    
End Function

Private Function IsDocumentReadable( _
        ByVal Path As String) As Boolean
        
    Dim extension As String
    extension = Fso.GetExtensionName(Path)
        
    Const allowedExtension As String = "doc"
        
    'Допускается doc или docx
    IsDocumentReadable = InStr(extension, allowedExtension) > vbEmpty

End Function

Private Function LoadFrom( _
        ByVal Path As String) As IWordDocument
        
    If (Fso.FileExists(Path) = False) Then _
       Exit Function
            
    Dim Document As IWordDocument
    Set Document = ApplicationWord.DocumentsOpen( _
                   FilePath:=Path, Visible:=False)
        
    Set LoadFrom = Document
        
End Function

Private Sub IWordDocumentsService_Dispose()
    
    If (ApplicationWord Is Nothing = True) Then _
       Exit Sub
    
    ApplicationWord.Quit
    
End Sub


