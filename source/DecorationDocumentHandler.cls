VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DecorationDocumentHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Comparison Results Decoration.Decoration Document Handler"
'@PredeclaredId
Option Explicit
Implements IDecorationDocumentHandler

Private Type TDecorationTableDocumentHandler
    Word As IWordDocumentsService
End Type

Private this As TDecorationTableDocumentHandler

Public Function Create( _
       ByVal Word As IWordDocumentsService) As IDecorationDocumentHandler
    
    With New DecorationDocumentHandler
    
        Set .Word = Word
        Set Create = .Self
        
    End With
    
End Function

Public Property Get Word() As IWordDocumentsService
    Set Word = this.Word
End Property
Public Property Set Word(ByVal RHS As IWordDocumentsService)
    Set this.Word = RHS
End Property

Public Property Get Self() As IDecorationDocumentHandler
    Set Self = Me
End Property

Private Function GetDocument( _
       ByVal Strategy As IComparisonTemplateFileStrategy, _
       ByVal Files As List_IFile_) As IWordDocument
    
    If (Strategy Is Nothing = True) Then _
        Exit Function

    If (Files Is Nothing = True) Then _
        Exit Function

    Dim DocumentPath As String

    Dim i As Long
    For i = 1 To Files.Count

        If Files(i).NameWithoutExtension = Strategy.TemplateFileMask Then _
            DocumentPath = Files(i).Path

    Next

    If DocumentPath = vbNullString Then _
        Exit Function

    If (Word Is Nothing = True) Then _
        Exit Function

    Dim Result As IWordDocument
    Set Result = Word.GetFrom(DocumentPath)

    Set GetDocument = Result
    
End Function
'---------------------------------------------
Private Function IDecorationDocumentHandler_GetDocument( _
        ByVal Strategy As IComparisonTemplateFileStrategy, _
        ByVal Files As List_IFile_) As IWordDocument
    
    Set IDecorationDocumentHandler_GetDocument = GetDocument(Strategy, Files)
    
End Function


