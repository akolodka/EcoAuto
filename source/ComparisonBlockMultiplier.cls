VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ComparisonBlockMultiplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Comparison Results Decoration.Comparison Block Decoration.Block Multiplier"
'@PredeclaredId
Option Explicit
Implements IComparisonBlockMultiplier

Private Type TComparisonBlockMultiplier
    Document As IEditableWordDocument
End Type

Private this As TComparisonBlockMultiplier

Public Function Create( _
       ByVal Document As IWordDocument) As IComparisonBlockMultiplier

    With New ComparisonBlockMultiplier
    
        Set .Document = Document
        Set Create = .Self
    
    End With

End Function

Public Property Get Document() As IEditableWordDocument
    Set Document = this.Document
End Property
Public Property Set Document(ByVal RHS As IEditableWordDocument)
    Set this.Document = RHS
End Property

Public Property Get Self() As IComparisonBlockMultiplier
    Set Self = Me
End Property

Private Sub ExecuteMultiplying()
    
    If (Document Is Nothing = True) Then _
       Exit Sub
        
    Dim Source As IEditableWordRange
    Set Source = Document.BookmarkedRange("template")
    
    Dim Destination As IEditableWordRange
    Set Destination = Document.StartOfDocument
    
    Source.Copy
    Destination.PasteAndFormat

End Sub
Private Sub Dispose()
    
    If (Document Is Nothing = True) Then _
        Exit Sub
    
    Dim Source As IEditableWordRange
    Set Source = Document.BookmarkedRange("template")

    Source.ClearContent
    Document.ClearEndParagraphs
    
End Sub
'---------------------------------------------
Private Sub IComparisonBlockMultiplier_ExecuteMultiplication()
    ExecuteMultiplying
End Sub
Private Sub IComparisonBlockMultiplier_Dispose()
    Dispose
End Sub

