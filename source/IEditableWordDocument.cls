VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IEditableWordDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("eco_transferToWord.Core.Word.Document")
'@Interface
Option Explicit

Public Function Range() As IWordRange
End Function

Public Function StartOfDocument() As IWordRange
End Function

Public Function BookmarkedRange(ByVal BookmarkName As String) As IWordRange
End Function

Public Sub ClearEndParagraphs()
End Sub


