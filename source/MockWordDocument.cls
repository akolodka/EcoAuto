VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MockWordDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "Tests.Mocks.Word.Document"
'@PredeclaredId
Option Explicit

Implements IWordDocument
Implements ISavableWordDocument
Implements IEditableWordDocument

Private Type TMockWordDocument
    WdRange As IWordRange
End Type

Private this As TMockWordDocument

Public Function Create() As IWordDocument
    
    With New MockWordDocument
        'В моке range создаётся один раз, впоследствии мок range.find внутри имеет счётчик поиска.
        Set .WdRange = MockWordRange.Create()
        Set Create = .Self
        
    End With
    
End Function

Public Property Get WdRange() As IWordRange
    Set WdRange = this.WdRange
End Property

Public Property Set WdRange(ByVal RHS As IWordRange)
    Set this.WdRange = RHS
End Property

Public Property Get Self() As IWordDocument
    Set Self = Me
End Property

Private Sub IEditableWordDocument_ClearEndParagraphs()
    
Debug.Print
Debug.Print "Mock WordDocument.ClearEndParagraphs"

End Sub

'---------------------------------------------
Private Property Get IWordDocument_ParentFolderPath() As String
    IWordDocument_ParentFolderPath = "...\ParentFolderPathOfMockWordDocument"
End Property

Private Property Get IWordDocument_Name() As String
    IWordDocument_Name = "MockWordDocument"
End Property

Private Property Get IWordDocument_Path() As String
    IWordDocument_Path = "...\MockWordDocument.docx"
End Property

Private Function IWordDocument_EndOfDocument() As IWordRange
    
End Function
'---------------------------------------------
Private Sub ISavableWordDocument_Save()
    
Debug.Print
Debug.Print "Mock WordDocument.Save"

End Sub

Private Sub ISavableWordDocument_CloseWithoutSave()

Debug.Print
Debug.Print "MockWordDocument.CloseWithoutSave"

End Sub

Private Sub ISavableWordDocument_SaveAs(ByVal FilePath As String)

Debug.Print
Debug.Print "Mock WordDocument.SaveAs " & vbNewLine & _
            "    FilePath: "; FilePath
End Sub
'---------------------------------------------
Private Function IEditableWordDocument_Range() As IWordRange
    Set IEditableWordDocument_Range = WdRange
End Function

Private Function IEditableWordDocument_EndOfDocument() As IWordRange
    Set IWordDocument_EndOfDocument = WdRange
End Function

Private Function IEditableWordDocument_StartOfDocument() As IWordRange
    Set IEditableWordDocument_StartOfDocument = WdRange
End Function

Private Function IEditableWordDocument_BookmarkedRange(ByVal BookmarkName As String) As IWordRange

    If (BookmarkName = vbNullString) Then _
        Exit Function
        
    Set IEditableWordDocument_BookmarkedRange = WdRange
    
End Function

Private Function IEditableWordDocument_DeleteBookmark(ByVal BookmarkName As String) As Variant

Debug.Print
Debug.Print "MockWordDocument.DeleteBookmark" & vbNewLine & _
            "   BookmarkName: " & BookmarkName
            
End Function
