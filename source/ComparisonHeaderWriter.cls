VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ComparisonHeaderWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder "eco_transferToWord.Core.Comparison Results Decoration.Comparison Block Decoration.Header Writer"
Option Explicit
Implements IComparisonHeaderWriter

Private Type TComparisonHeaderWriter
    Writer As IDocumentWriter
End Type

Private this As TComparisonHeaderWriter

Public Function Create( _
       ByVal Writer As IDocumentWriter) As IComparisonHeaderWriter
    
    With New ComparisonHeaderWriter
        
        Set .Writer = Writer
        Set Create = .Self
        
    End With
    
End Function

Public Property Get Self() As IComparisonHeaderWriter
    Set Self = Me
End Property

Public Property Get Writer() As IDocumentWriter
    Set Writer = this.Writer
End Property
Public Property Set Writer(ByVal RHS As IDocumentWriter)
    Set this.Writer = RHS
End Property

Private Sub FillBlockHeaders( _
       ByVal Model As IComparisonResultHeaderModel)
       
    Writer.InsertIntoDocument Model.Name, InsertType.ReplaceFirstEntry
    Writer.InsertIntoDocument Model.Subname, InsertType.ReplaceFirstEntry
        
    Writer.InsertIntoDocument Model.Unit, InsertType.ReplaceTwice
    
End Sub
'---------------------------------------------
Private Sub IComparisonHeaderWriter_FillBlockHeaders(ByVal Model As IComparisonResultHeaderModel)
    FillBlockHeaders Model
End Sub

