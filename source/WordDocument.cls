VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WordDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("eco_transferToWord.Core.Word.Document")
'@PredeclaredId
Option Explicit

Implements IWordDocument
Implements ISavableWordDocument
Implements IEditableWordDocument

Private Type TWordDocument
    Document As Word.Document
End Type

Private this As TWordDocument

Public Function Create( _
       ByVal Document As Word.Document) As IWordDocument
    
    With New WordDocument
        
        Set .Document = Document
        Set Create = .Self
        
    End With
    
End Function

Public Property Get Document() As Word.Document
    Set Document = this.Document
End Property
Public Property Set Document(ByVal RHS As Word.Document)
    Set this.Document = RHS
End Property

Public Property Get Self() As IWordDocument
    Set Self = Me
End Property

Public Function StartOfDocument() As IWordRange

    Dim startPosition As Long
    startPosition = Document.Range.Start
    
    Dim Result As IWordRange
    Set Result = WordRange.Create(Document.Range(startPosition, startPosition))
    
    Set StartOfDocument = Result

End Function

Private Function BookmarkedRange( _
       ByVal BookmarkName As String) As IWordRange
    
    If (BookmarkName = vbNullString) Then _
        Exit Function
    
    If (Document.Bookmarks.Exists(BookmarkName) = False) Then _
        Exit Function
    
    Dim TargetRange As Word.Range
    Set TargetRange = Document.Bookmarks(BookmarkName).Range
    
    Dim Result As IWordRange
    Set Result = WordRange.Create(TargetRange)
    
    Set BookmarkedRange = Result
    
End Function
Private Function ClearEndParagraphs()
    Document.Range(Document.Range.End - 2).Delete
End Function
'---------------------------------------------
Private Property Get IWordDocument_ParentFolderPath() As String
    IWordDocument_ParentFolderPath = Document.Path
End Property
Private Property Get IWordDocument_Name() As String
    IWordDocument_Name = Document.Name
End Property
Private Property Get IWordDocument_Path() As String
    IWordDocument_Path = Document.FullName
End Property
'---------------------------------------------
Private Sub ISavableWordDocument_Save()
    Document.Save
End Sub
Private Sub ISavableWordDocument_CloseWithoutSave()
    Document.Close SaveChanges:=False
End Sub
Private Sub ISavableWordDocument_SaveAs(ByVal FilePath As String)
    Document.SaveAs2 FilePath
End Sub
'---------------------------------------------
Private Function IEditableWordDocument_Range() As IWordRange
    
    'Обязательно новый rangе каждый раз, старый range затирается
    Set IEditableWordDocument_Range = WordRange.Create(Document.Range)
    
End Function
Private Function IEditableWordDocument_StartOfDocument() As IWordRange
    Set IEditableWordDocument_StartOfDocument = StartOfDocument
End Function
Private Function IEditableWordDocument_BookmarkedRange(ByVal BookmarkName As String) As IWordRange
    Set IEditableWordDocument_BookmarkedRange = BookmarkedRange(BookmarkName)
End Function
Private Sub IEditableWordDocument_ClearEndParagraphs()
    ClearEndParagraphs
End Sub
