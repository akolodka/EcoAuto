VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ResultDocumentTemplateLoader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Template Loader"
'@PredeclaredId
Option Explicit
Implements IResultDocumentTemplateLoader

Private Type TTemplateLoader

    Fso As IFileSystemObject
    Word As IWordDocumentsService

    TemporaryFolderPath As String

End Type

Private this As TTemplateLoader

Public Function Create( _
        ByVal Initial As ITemplateLoaderInitialization) As IResultDocumentTemplateLoader
    
    With New ResultDocumentTemplateLoader
        
        Set .Fso = Initial.Fso
        Set .Word = Initial.Word
        
        .TemporaryFolderPath = Initial.TemporaryFolderPath
        
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

Public Property Get TemporaryFolderPath() As String
    TemporaryFolderPath = this.TemporaryFolderPath
End Property
Public Property Let TemporaryFolderPath(ByVal RHS As String)
    this.TemporaryFolderPath = RHS
End Property

Public Property Get Self() As IResultDocumentTemplateLoader
    Set Self = Me
End Property

Private Function GetTemplates() As List_IWordDocument_
    
    Dim Result As List_IWordDocument_
    Set Result = New List_IWordDocument_
    
    Dim Files As List_IFile_
    Set Files = Fso.GetFolderFiles(TemporaryFolderPath)

    Dim i As Long
    For i = 1 To Files.Count

        Dim doc As IWordDocument
        Set doc = Word.GetFrom(Files(i).Path)

        If (doc Is Nothing = False) Then _
           Result.Add doc
            
    Next
    
    Set GetTemplates = Result

End Function
'---------------------------------------------
Private Function IResultDocumentTemplateLoader_GetTemplates() As List_IWordDocument_
    Set IResultDocumentTemplateLoader_GetTemplates = GetTemplates
End Function


