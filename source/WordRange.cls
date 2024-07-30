VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WordRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("eco_transferToWord.Core.Word.Range")
'@PredeclaredId
Option Explicit

Implements IWordRange
Implements IFindableWordRange
Implements IEditableWordRange
Implements IUpdatableWordRange

Private Type TWordRange
    DocumentRange As Word.Range
End Type

Private this As TWordRange

Public Function Create( _
       ByVal DocumentRange As Word.Range) As IWordRange
    
    With New WordRange
        
        Set .DocumentRange = DocumentRange
        Set Create = .Self
        
    End With
    
End Function

Public Property Get Self() As IWordRange
    Set Self = Me
End Property

Public Property Get DocumentRange() As Word.Range
    Set DocumentRange = this.DocumentRange
End Property
Public Property Set DocumentRange(ByVal RHS As Word.Range)
    Set this.DocumentRange = RHS
End Property

Public Property Get LastTable() As IFormattableWordTable
    
    Dim Table As Word.Table
    Set Table = DocumentRange.Tables(DocumentRange.Tables.Count)
    
    Dim Result As IFormattableWordTable
    Set Result = WordTable.Create(Table)
    
    Set LastTable = Result
    
End Property

Public Property Get Tables() As List_IWordTable_
    
    Dim Result As List_IWordTable_
    Set Result = New List_IWordTable_
    
    Dim Table As Word.Table
    For Each Table In DocumentRange.Tables
    
        Result.Add _
        WordTable.Create(Table)
        
    Next
    
    Set Tables = Result

End Property
Public Function FindExecute( _
        Optional ByVal FindText As String) As Boolean
    
    If (DocumentRange Is Nothing = True) Then _
        Exit Function
    
    FindExecute = DocumentRange.Find.Execute(FindText)

End Function

Public Sub Copy()
    
    If (DocumentRange Is Nothing = True) Then _
        Exit Sub
        
    DocumentRange.Copy

End Sub

Public Sub PasteAndFormat()
      
    If (DocumentRange Is Nothing = True) Then _
        Exit Sub
      
    DocumentRange.PasteAndFormat wdPasteDefault
    
End Sub

Public Sub PasteAsText( _
       ByVal Text As String)

    DocumentRange.Text = Text
    DocumentRange.Font.ColorIndex = wdBlack
    DocumentRange.HighlightColorIndex = wdNoHighlight

End Sub

Public Sub ClearContent()
    DocumentRange.Text = vbNullString
End Sub

Public Sub UpdateFields()

    If (DocumentRange Is Nothing = True) Then _
        Exit Sub
    
    DocumentRange.Fields.Update
    
End Sub
'---------------------------------------------
Private Property Get IWordRange_LastTable() As IFormattableWordTable
    Set IWordRange_LastTable = LastTable
End Property
Private Property Get IWordRange_Tables() As List_IWordTable_
    Set IWordRange_Tables = Tables
End Property
'---------------------------------------------
Private Function IFindableWordRange_FindExecute(Optional ByVal FindText As String) As Boolean
    IFindableWordRange_FindExecute = FindExecute(FindText)
End Function
'---------------------------------------------
Private Sub IEditableWordRange_Copy()
    Copy
End Sub
Private Sub IEditableWordRange_PasteAndFormat()
    PasteAndFormat
End Sub
Private Sub IEditableWordRange_PasteAsText(ByVal Text As String)
    PasteAsText Text
End Sub
Private Sub IEditableWordRange_ClearContent()
    ClearContent
End Sub
'---------------------------------------------
Private Sub IUpdatableWordRange_UpdateFields()
    UpdateFields
End Sub


